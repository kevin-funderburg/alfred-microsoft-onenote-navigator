#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
import urllib2
import time
import argparse
from plistlib import readPlist
import sqlite3
from sqlite3 import Error
from notebook_item import NotebookItem
import queries

from workflow import (Workflow3, ICON_INFO, ICON_WARNING,
                      ICON_ERROR, MATCH_ALL, MATCH_ALLCHARS,
                      MATCH_SUBSTRING, MATCH_STARTSWITH)
from workflow.util import run_trigger, run_command, unicodify, utf8ify, unset_config
from workflow.background import run_in_background, is_running

__version__ = '1.3.1'

wf = None  # type: Workflow3
log = None
# sub = []
# subtitle = ""
# url = ""
# urlbase = None
all_data = []
nitems =[]
path_map = {}
# path_tablee = [][]

HELP_URL = 'https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator'

# GitHub repo for self-updating
UPDATE_SETTINGS = {'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator'}
DEFAULT_SETTINGS = {'urlbase': None}

ICON_ONENOTE = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"

ONENOTE_APP_SUPPORT = "~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support"
ONENOTE_FULL_SEARCH_PATH = ONENOTE_APP_SUPPORT + "/Microsoft User Data/OneNote/15.0/FullTextSearchIndex/"
ONENOTE_USER_INFO_CACHE = ONENOTE_APP_SUPPORT + "/Microsoft/UserInfoCache/9478a1a4ec3795b7_LiveId.db"
ONENOTE_PLIST_PATH = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
ONENOTE_USER_UID = None
ONENOTE_PLIST = None
ALL_DB_PATHS = []

MERGED_DB = 'merged-onenote-data.db'

data = []
sec_paths = []


def search_all_db_entries():
    conn = create_connection(wf.datafile(MERGED_DB))

    log.debug('\nstarting loop')
    for row in conn.execute(str("SELECT * FROM Entities;")):
        item = NotebookItem(row)
        item.set_path(get_page_path(item))
        it = wf.add_item(
                item.Title,
                arg=item.GUID,
                subtitle=item.path,
                # arg=dict(row),
                # arg=make_url(row),
                uid=item.GUID,
                autocomplete=item.Title,
                valid=True,
                icon=item.icon,
                icontype="file"
            )
        data.append(it)
    log.debug('\nloop complete')
    wf.send_feedback()


def get_notebook_items(args):

    query = populate_query(args)
    # if search == 'recent':
    #     query = "SELECT * FROM Entities ORDER BY RecentTime DESC;"
    # elif search == 'modified':
    #     query = "SELECT * FROM Entities ORDER BY LastModifiedTime DESC;"
    # elif search == 'all':
    #     query = "SELECT * FROM Entities;"

    # conn = create_connection(wf.datafile(MERGED_DB))
    #
    # for row in conn.execute(str(query)):
    for row in run_query(query):
        ni = NotebookItem(row)
        ni.set_path(get_page_path(ni))
        nitems.append(ni)
        it = wf.add_item(
                ni.Title,
                arg=ni.GUID,
                subtitle=ni.path,
                uid=ni.GUID,
                autocomplete=ni.Title,
                valid=True,
                icon=ni.icon,
                icontype="file"
            )
    log.info('loop complete')
    wf.send_feedback()
    return nitems


def update_path_map():
    conn = create_connection(wf.datafile(MERGED_DB))
    query = "SELECT * FROM Entities;"

    for row in conn.execute(str(query)):
        ni = NotebookItem(row)
        ni.set_path(get_page_path(ni))


def make_url(item):
    # log.debug('row type is: ' + str(typ/e(row)))
    # assert isinstance(row, list), 'expected <list> and got ' + str(type(row))

    conn = create_connection(wf.datafile(MERGED_DB))
    cur = conn.cursor()
    set_user_uid(None)
    base = 'onenote:https://d.docs.live.net/{0}/Documents/'.format(ONENOTE_USER_UID)
    path = get_page_path(item)
    url = "{0}{1}".format(base, path)

    if item.Type <= 2:   # if section or page
        # cur.execute(str("SELECT * FROM Entities WHERE GOID = \"{0}\"".format(item.ParentGOID)))
        # r = cur.fetchone()
        section_id = item.GUID
        if item.Type == 2:
            url = "{0}&section-id={1}&end".format(url, section_id)
        else:
            url = "{0}&section-id={1}&page-id={2}&end".format(url, section_id, item.GUID)

    return url


def get_page_path(ni):
    global path_map

    if ni.GOID not in path_map:
        if ni.Type == 4:
            entry = {ni.GOID: ni.Title}
            path_map.update(entry)
            return path_map[ni.GOID]

        parent_path = None
        if ni.Type == 3:
            if ni.has_grandparent():
                if ni.last_grandparent in path_map:
                    parent_path = path_map[ni.last_grandparent]
                else:
                    parent_row = NotebookItem(get_parent_row(ni.last_grandparent))
                    parent_path = get_page_path(parent_row)
            else:
                parent_path = get_parent_path(ni)

        elif ni.Type == 2 or ni.Type == 1:
            parent_path = get_parent_path(ni)

        page_path = "{0}/{1}".format(parent_path, ni.Title)
        entry = {ni.GOID: page_path}
        path_map.update(entry)

    return path_map[ni.GOID]


def get_parent_row(parent_goid):
    conn = create_connection(wf.datafile(MERGED_DB))
    for row in conn.execute(str("SELECT * FROM Entities WHERE GOID = \"{0}\"".format(parent_goid))):
        return row


def get_parent_path(item):
    global path_map
    if item.ParentGOID in path_map:
        return path_map[item.ParentGOID]
    else:
        parent_row = NotebookItem(get_parent_row(item.ParentGOID))
        return get_page_path(parent_row)


def get_row(uid):
    sql = "SELECT * FROM Entities WHERE GUID = \"{0}\"".format(uid)
    conn = create_connection(wf.datafile(MERGED_DB))
    cur = conn.cursor()
    cur.execute(sql)
    res = cur.fetchall()
    if res:
        log.debug('row type in get_row: ' + str(type(res)))
        return res[0]
    else:
        return None


def get_children(sec_guid):
    return run_query(queries.get_children(sec_guid))


def create_db():
    sql = queries.create_merged_db()
    conn = create_connection(wf.datafile(MERGED_DB))
    cur = conn.cursor()
    lines = sql.splitlines()
    for line in lines:
        if line is not "":
            cur.execute(str(line))

    conn.commit()
    rows = cur.fetchall()
    for r in rows:
        print(r[0])


def update_db(): create_db()


def reset_db():
    run_query(queries.reset_db())


def get_page_name(GUID):
    conn = create_connection(wf.datafile(MERGED_DB))
    conn.text_factory = unicode
    cur = conn.cursor()
    cur.execute(str("SELECT * FROM Entities WHERE GUID = \"{0}\"".format(GUID)))
    r = cur.fetchone()
    return r


def populate_query(args):
    sql = None
    if args.all:
        log.debug('Searching all')
        sql = queries.get_all_items()
    elif args.recent:
        log.debug('Searching recent items')
        sql = queries.get_recent_items()
    elif args.modified:
        log.debug('Searching modified items')
        sql = queries.get_last_modified()
    elif args.update:
        log.debug('Updating merged db')
        sql = queries.create_merged_db()

    return sql


def run_query(sql):
    db_path = wf.datafile(MERGED_DB)

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    log.debug(sql)
    cursor.execute(sql)
    results = cursor.fetchall()
    log.debug("Found {0} results".format(len(results)))
    cursor.close()
    return results


def create_connection(db_file):
    """ create a database connection to the SQLite database
        specified by the db_file
    :param db_file: database file
    :return: Connection object or None
    """
    conn = None
    rdb = r"{0}".format(os.path.expanduser(db_file))
    try:
        conn = sqlite3.connect(rdb)
        conn.row_factory = sqlite3.Row
        conn.text_factory = unicode
    except Error as e:
        raise e

    return conn


def set_user_uid(path):
    try:
        global ONENOTE_USER_UID
        ONENOTE_USER_UID = re.search('.*UserInfoCache/(.*)_LiveId\\.db', ONENOTE_USER_INFO_CACHE).group(1)
    except Error as e:
        raise Exception("Unable to create the OneNote user name")


def get_results(sql):
    wf.logger.debug(sql)
    results = run_query(sql)

    if not results:
        wf.add_item('No items', icon=ICON_WARNING)
    else:
        for result in results:
            item = NotebookItem(result)
            item.set_path(get_page_path(item))
            nitems.append(item)
            log.debug(item)
            it = wf.add_item(
                title=item.Title,
                arg=item.GUID,
                subtitle=item.path,
                uid=item.GUID,
                autocomplete=item.Title,
                valid=True,
                icon=item.icon,
                icontype="file"
            )


def parse_args():
    parser = argparse.ArgumentParser(description="Search OneNote")

    parser.add_argument('--seturl', dest='urlbase', nargs='?', default=None)
    parser.add_argument('--type', dest='type', nargs='?', default=None)
    parser.add_argument('--searchall', dest='searchall', action='store_true')
    parser.add_argument('-a', '--all', dest='all', action='store_true',
                        help='search unordered notebook items')
    parser.add_argument('-r', '--recent', dest='recent', action='store_true',
                        help='search recent notebook items')
    parser.add_argument('-m', '--modified', dest='modified', action='store_true',
                        help='search last modified notebook items')
    parser.add_argument('-u', '--update', dest='update', action='store_true',
                        help='update consolidated database')
    parser.add_argument('-o', '--open', dest='open', nargs='?', default=None,
                        help='open the notebook item with provided guid')
    parser.add_argument('-b', '--browse', dest='browse', nargs='?', default=None,
                        help='browse notebooks from top level')

    parser.add_argument('-n', dest='name', nargs='?', default=None)
    parser.add_argument('--url', dest='url', nargs='?', default=None)
    parser.add_argument('--warn', dest='warn', nargs='?', default=None)
    parser.add_argument('query', nargs='?', default=None)
    log.debug(wf.args)
    args = parser.parse_args(wf.args)
    return args


def main(wf):
    log.debug('Started workflow')
    args = parse_args()

    global path_map

    if wf.update_available:
        # Add a notification to top of Script Filter results
        wf.add_item('New version available',
                    'Action this item to install the update',
                    autocomplete='workflow:update',
                    icon=ICON_INFO)

    if args.open:
        ni = NotebookItem(get_row(args.open))

        if ni.Type == 1:
            log.debug('argument is a page, preparing to open...' + ni.Title)
            open_url(ni.url)
            # open_url(make_url(item))
        else:
            log.info("setting 'q' to: {0}".format(args.open))
            log.debug('argument is not a page, gathering sub-pages...')
            # get_section_pages(args.open)
            wf.setvar("q", ni.GUID, True)
            run_trigger('browse')
        return 0

    query = populate_query(args)
    get_results(query)
    wf.send_feedback()
    return


    if args.browse:
        get_children(args.browse)
        return 0

    if args.name:
        get_children(args.name)
        return 0

    if args.open:
        ni = NotebookItem(get_row(args.open))

        if ni.Type == 1:
            log.debug('argument is a page, preparing to open...' + ni.Title)
            open_url(ni.url)
            # open_url(make_url(item))
        else:
            log.info("setting 'q' to: {0}".format(args.open))
            log.debug('argument is not a page, gathering sub-pages...')
            # get_section_pages(args.open)
            wf.setvar("q", ni.GUID, True)
            run_trigger('browse')
        return 0

    if args.all:
        get_notebook_items('all')
        return 0

    if args.recent:
        get_notebook_items('recent')
        return 0

    if args.modified:
        update_db()
        get_notebook_items('modified')
        return 0

    if args.update:
        update_db()
        wf.add_item('Updating OneNote cache.',
                    'Give me a minute and then try searching again.',
                    valid=False,
                    icon=ICON_INFO)
        wf.send_feedback()
        return 0






def key_for_data(data): return '{}'.format(data['Name'])


def get_notebook_data():
    """
    get the data for only the notebooks,
    not their subsections

    :return: dict
    """
    data = []
    for n in ONENOTE_PLIST:
        sub = "{0}".format(n["Name"])
        log.debug("sub is: " + sub)
        url = makeurl(sub)
        item = {
            "Name": n["Name"],
            "subtitle": sub,
            "arg": url,
            "autocomplete": n["Name"],
            "icon": "icons/notebook.png"
        }
        data.append(item)
    return data


def get_all_data(node, prefix):
    """ recursively get every section of the parent

    this saves all the data in the global variable all_data()

    :param node: entry point of search
    :param prefix: subtitle of page, pass as None for root
    :return: none

    """
    global subtitle
    global url
    global all_data
    global gosids

    if len(node) > 0 and "Name" not in node:
        for n in node:
            get_all_data(n, prefix)
        prefix = None
        url = ""
        gosids = []

    if "Name" in node:
        if "Children" in node:
            gosids.append(node["Gosid"])

            if prefix is None:
                pre = node["Name"]
                url = "{0}{1}".format(urlbase, pre)
                subtitle = "{0}".format(node["Name"])
                data = {
                    "Name": node["Name"],
                    "subtitle": subtitle,
                    "arg": subtitle,
                    "autocomplete": node["Name"],
                    "icon": "icons/notebook.png",
                    "url": url,
                    "gosids": gosids
                }
                all_data.append(data)

            else:
                pre = prefix + " > " + node["Name"]
                subtitle = pre
                url = makeurl(subtitle)
                data = {
                    "Name": node["Name"],
                    "subtitle": pre,
                    "arg": subtitle,
                    "autocomplete": node["Name"],
                    "icon": "icons/section.png",
                    "url": url,
                    "gosids": gosids
                }
                all_data.append(data)

            get_all_data(node["Children"], pre)

        else:
            subtitle = "{0} > {1}".format(prefix, node["Name"])
            url = makeurl(subtitle)
            data = {
                "Name": node["Name"],
                "subtitle": subtitle,
                "arg": url + ".one",
                "autocomplete": node["Name"],
                "icon": ICON_PAGE,
                "gosids": gosids
            }
            # pages =
            all_data.append(data)


def get_pages(GOSID):
    query = "SELECT	*" \
            "FROM	Entities" \
            "WHERE 	ParentGOID = (" \
                "SELECT	GOID" \
                "FROM	Entities" \
                "WHERE	GOSID = '{0}'" \
            ");".format(GOSID)
    conn = create_connection(wf.datafile(MERGED_DB))
    for row in conn.execute(query):
        data = {
            "Name": row[str('Title')],
            "subtitle": subtitle,
            "arg": url + ".one",
            "autocomplete": node["Name"],
            "icon": ICON_PAGE,
            "gosids": gosids
        }


def makeurl(prefix):
    newurl = "{0}{1}".format(urlbase, prefix.replace(" > ", "/"))
    return encode_url(newurl)


def get_child(childstr):
    """ finds the child of the onenote plist

    :param childstr: path of child, formatted as '[element] > [element]'
    :return: child element of onenote plist
    """
    items = childstr.split(" > ")
    child = None
    for x in range(len(items)):
        if x == 0:
            for p in ONENOTE_PLIST:
                if p["Name"] == items[0]:
                    child = p["Children"]
                    break
        else:
            for c in child:
                if c["Name"] == items[x]:
                    if "Children" in c:
                        child = c["Children"]
                    break
    return child


def browse_child():
    """ get the children of the defined element

    element is defined by the alfred variable q
    which looks like [element] > [element]
    :return: dict
    """
    data = []
    q = os.getenv('q')
    log.info("q is: {0}".format(q))
    child = get_child(q)
    for c in child:
        sub = "{0} > {1}".format(q, c["Name"])
        url = makeurl(sub)
        if "Children" in c:
            item = {
                "Name": c["Name"],
                "subtitle": sub,
                "arg": sub,
                "autocomplete": c["Name"],
                "icon": "icons/section.png",
                "url": url
            }
        else:
            item = {
                "Name": c["Name"],
                "subtitle": sub,
                "arg": url + ".one",
                "autocomplete": c["Name"],
                "icon": "icons/page.png",
            }

        data.append(item)

    return data


def set_base_url(urlbase):
    if 'sharepoint' in urlbase:
        # sharepoint URLs will not work, pass as an error
        print("sharepoint")
        return 0
    elif 'onenote:https' in urlbase:
        # extract onenote url
        urlbase = re.search('(onenote:.*)', urlbase).group(1)
        url = urlbase.split("/")
        urlbase = url[0] + "//" + url[2] + "/" + url[3] + "/" + url[4] + "/"
        # write to plist
        wf.setvar("urlbase", urlbase, True)
        # tell alfred how to display a success notification
        print("true")
    else:
        # tell alfred how to display alert that the url is bad
        print("invalid")
    return 0


def warn(query):
    if "sharepoint" in query:
        wf.add_item(title='URL is a Microsoft Sharepoint URL, these '
                          'cannot be opened locally.',
                    subtitle="OneNote must be in OneDrive for URLs to work",
                    valid=False,
                    icon=ICON_ERROR)
    else:
        wf.add_item(title='Argument is not a valid OneNote url.',
                    subtitle="Right click a OneNote notebook/section/page "
                             "& choose 'Copy Link to ...' & try again",
                    valid=False,
                    icon=ICON_WARNING)
    wf.send_feedback()


def browse_notebooks():
    """get the notebooks of the onenote plist

    :return: none
    """
    for n in ONENOTE_PLIST:
        sub = "{0}".format(n["Name"])
        url = makeurl(sub)
        it = wf.add_item(title=n["Name"],
                         subtitle=sub,
                         arg="",
                         autocomplete=n["Name"],
                         valid=True,
                         icon="icons/notebook.png",
                         icontype="file",
                         quicklookurl=ICON_ONENOTE)
        it.add_modifier('cmd', subtitle="open in OneNote", arg=url, valid=True)
        it.setvar('theTitle', n["Name"])
        it.setvar("q", sub)


def encode_url(url):
    """
    encode url for shell to open
    :rtype: str
    """
    delims = {
        " ": "%20",
        "&": "%26",
        "{": "%7B",
        "}": "%7D",
    }

    for key, val in delims.items():
        url = url.replace(key, val)
    return url


def open_url(url):
    url = encode_url(url)
    log.debug('\n\nafter ut8ify: {0}\n\n'.format(url))
    run_command(['open', url])


def clear_config():
    unset_config('q')
    unset_config('theTitle')
    log.info("'q' and 'theTitle' cleared")
    pass


def init_wf():
    global wf, log
    wf = Workflow3(default_settings=DEFAULT_SETTINGS,
                   update_settings=UPDATE_SETTINGS,
                   help_url=HELP_URL)
    log = wf.logger


if __name__ == "__main__":
    init_wf()
    sys.exit(wf.run(main))

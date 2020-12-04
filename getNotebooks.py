#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
import argparse
from plistlib import readPlist
import sqlite3
from sqlite3 import Error

from workflow import (Workflow3, ICON_INFO, ICON_WARNING,
                      ICON_ERROR, MATCH_ALL, MATCH_ALLCHARS,
                      MATCH_SUBSTRING, MATCH_STARTSWITH)
from workflow.util import run_trigger, run_command, unset_config

__version__ = '1.3.1'

wf = None
log = None
sub = []
subtitle = ""
url = ""
urlbase = None
gosids = []
all_data = []

HELP_URL = 'https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator'

# GitHub repo for self-updating
UPDATE_SETTINGS = {'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator'}
DEFAULT_SETTINGS = {'urlbase': None}

ONENOTE_APP_SUPPORT = "~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support"
ONENOTE_FULL_SEARCH_PATH = ONENOTE_APP_SUPPORT + "/Microsoft User Data/OneNote/15.0/FullTextSearchIndex/"
ONENOTE_USER_INFO_CACHE = ONENOTE_APP_SUPPORT + "/Microsoft/UserInfoCache/9478a1a4ec3795b7_LiveId.db"
ONENOTE_PLIST_PATH = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
ONENOTE_ICON = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"
ONENOTE_USER_UID = None
ONENOTE_PLIST = None  # type: dict
ALL_DB_PATHS = []


def db_test():
    set_user_uid(ONENOTE_USER_INFO_CACHE)

    i = 0
    union_all = ""
    for f in os.listdir(os.path.expanduser(ONENOTE_FULL_SEARCH_PATH)):
        if '.db' in f and 'journal' not in f:
            path = ONENOTE_FULL_SEARCH_PATH + f
            ALL_DB_PATHS.append(ONENOTE_FULL_SEARCH_PATH + f)

    attach_dbs = ""
    db = ALL_DB_PATHS[0]

    # create a database connection
    # conn.row_factory = sqlite3.Row
    conn = create_connection(db)
    cur = conn.cursor()

    cur.execute(union_all)
    # print(attach_dbs)
    # print(union_all)
    rows = cur.fetchall()
    for r in rows:
        print(r[0])


def create_db():
    sql = make_sql_script()
    conn = create_connection('all-onenote.db')
    cur = conn.cursor()
    lines = sql.splitlines()
    for line in lines:
        if line is not "":
            cur.execute(str(line))

    conn.commit()
    rows = cur.fetchall()
    for r in rows:
        print(r[0])


def reset_db():
    sql = "DROP TABLE Entities;\n"


def make_sql_script():
    for f in os.listdir(os.path.expanduser(ONENOTE_FULL_SEARCH_PATH)):
        if '.db' in f and 'journal' not in f:
            ALL_DB_PATHS.append(ONENOTE_FULL_SEARCH_PATH + f)

    drop_table = "DROP TABLE IF EXISTS Entities;\n"

    create_table = "CREATE TABLE Entities (" \
                   "Type INTEGER, " \
                   "GOID NVARCHAR(50) NOT NULL, " \
                   "GUID NVARCHAR(38) NOT NULL, " \
                   "GOSID NVARCHAR(50), " \
                   "ParentGOID NVARCHAR(50), " \
                   "GrandparentGOIDs TEXT, " \
                   "ContentRID NVARCHAR(50), " \
                   "RootRevGenCount INTEGER, " \
                   "LastModifiedTime INTEGER, " \
                   "RecentTime INTEGER, " \
                   "PinTime INTEGER, " \
                   "Color INTEGER, " \
                   "Title TEXT, " \
                   "EnterpriseIdentity TEXT" \
                   ")"

    assert (len(ALL_DB_PATHS) > 0)
    attaches = ""
    inserts = ""
    keys = "Type, GOID, GUID, GOSID, ParentGOID, GrandparentGOIDs, " \
           "ContentRID, RootRevGenCount, LastModifiedTime, RecentTime, " \
           "PinTime, Color, Title, EnterpriseIdentity"

    for i in range(1, len(ALL_DB_PATHS)):
        attaches += "ATTACH DATABASE \"{0}\" as db{1};\n".format(os.path.expanduser(ALL_DB_PATHS[i]), i)
        inserts += "INSERT INTO Entities SELECT {0} FROM db{1}.Entities;\n".format(keys, i)

    sql = drop_table\
          + create_table + \
          "\n\n" + \
          attaches + \
          "\n\n" \
          + inserts
    # print(create_table + "\n\n" + attaches + "\n\n" + inserts)

    return sql


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
    except Error as e:
        print(e)

    return conn


def set_user_uid(path):
    try:
        global ONENOTE_USER_UID
        ONENOTE_USER_UID = re.search('.*UserInfoCache/(.*)_LiveId\\.db', ONENOTE_USER_INFO_CACHE).group(1)
    except Error as e:
        raise Exception("Unable to create the OneNote user name")


def main(wf):
    global urlbase
    global ONENOTE_PLIST_PATH
    global ONENOTE_PLIST

    if wf.update_available:
        # Add a notification to top of Script Filter results
        wf.add_item('New version available',
                    'Action this item to install the update',
                    autocomplete='workflow:update',
                    icon=ICON_INFO)

    parser = argparse.ArgumentParser()
    # --seturl argument and save its value to 'urlbase' (dest).
    parser.add_argument('--seturl', dest='urlbase', nargs='?', default=None)
    parser.add_argument('--browse', dest='browse', nargs='?', default=None)
    parser.add_argument('--type', dest='type', nargs='?', default=None)
    parser.add_argument('--searchall', dest='searchall', action='store_true')
    parser.add_argument('--db', dest='db', action='store_true')
    parser.add_argument('--warn', dest='warn', nargs='?', default=None)
    parser.add_argument('--url', dest='url', nargs='?', default=None)
    parser.add_argument('query', nargs='?', default=None)
    args = parser.parse_args(wf.args)

    ####################################################################
    # Save the provided URL base
    ####################################################################
    if args.urlbase:  # Script was passed as a URL
        set_base_url(args.urlbase)
        return 0

    ####################################################################
    # Warn if bad url was passed in
    ####################################################################
    if args.warn:
        warn(args.warn)
        return 0

    ####################################################################
    # Open OneNote URL or store variables & make recursive call
    ####################################################################
    if args.url:
        if 'onenote:https' in args.url:
            open_url(args.url)
        else:
            log.info("setting 'q' to: {0}".format(args.url))
            wf.setvar("q", args.url, True)
            wf.setvar("theTitle", os.getenv('theTitle'), True)
            run_trigger('browse_section', wf.bundleid)
        return 0

    ####################################################################
    # Check that we have a base URL saved
    ####################################################################
    log.debug('environmental var urlbase is: {0}'.format(os.getenv('urlbase')))
    if os.getenv('urlbase') == "":
        wf.add_item('No URL set yet.',
                    'Please use seturl to set your OneNote url base.',
                    valid=False,
                    icon=ICON_WARNING)
        wf.send_feedback()
        return 0

    try:
        ONENOTE_PLIST_PATH = os.path.expanduser(ONENOTE_PLIST_PATH)
    except OSError:
        wf.add_item('OneNote plist not found.',
                    'Make sure OneNote is installed correctly and try again.',
                    valid=False,
                    icon=ICON_WARNING)
        wf.send_feedback()
        return 0

    # get plist data from OneNote plist
    ONENOTE_PLIST = readPlist(ONENOTE_PLIST_PATH)
    urlbase = os.getenv('urlbase')

    if args.type == 'searchall':
        clear_config()
        get_all_data(ONENOTE_PLIST, None)

        items = wf.filter(args.query,
                          all_data,
                          key_for_data,
                          min_score=30)
                          # match_on=MATCH_STARTSWITH | MATCH_SUBSTRING)
                          # match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING)

        if not items:
            wf.add_item(title='No matches',
                        icon=ICON_WARNING)

        for item in items:
            it = wf.add_item(
                item["Name"],
                item["subtitle"],
                arg=item["arg"],
                uid=item['subtitle'],
                autocomplete=item["Name"],
                valid=True,
                icon=item["icon"],
                icontype="file"
            )
            if "onenote:https" not in item["arg"]:
                it.add_modifier(
                    'cmd',
                    subtitle="open in OneNote",
                    arg=item["url"] + ".one",
                    valid=True
                )
                it.setvar('theTitle', item["Name"])

        wf.send_feedback()
        return 0

    elif args.browse:
        log.debug("browse is: {0}".format(args.browse))
        log.debug("query is: {0}".format(args.query))

        if args.browse == 'notebooks':
            data = get_notebook_data()
            items = wf.filter(
                args.query,
                data,
                key_for_data,
                min_score=30,
                match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING
            )
            if not items:
                wf.add_item(title="No notebooks matching '" + args.query + "'",
                            icon=ICON_WARNING)

            for item in items:
                it = wf.add_item(
                    title=item["Name"],
                    subtitle=item["subtitle"],
                    arg="",
                    uid=item["Name"],
                    autocomplete=item["Name"],
                    valid=True,
                    icon=item["icon"],
                    icontype="file"
                )
                it.add_modifier('cmd',
                                subtitle="open in OneNote",
                                arg=item["arg"],
                                valid=True)
                it.setvar('theTitle', item["Name"])
                it.setvar("q", item["subtitle"])

            wf.send_feedback()
            return 0

        else:
            data = browse_child()
            items = wf.filter(args.query,
                              data,
                              key_for_data,
                              min_score=30,
                              # match_on=MATCH_STARTSWITH | MATCH_SUBSTRING)
                              match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING)

            if not items:
                wf.add_item(title='No matches',
                            icon=ICON_WARNING)

            for item in items:
                it = wf.add_item(
                    title=item["Name"],
                    subtitle=item["subtitle"],
                    arg=item["arg"],
                    uid=item["arg"],
                    autocomplete=item["Name"],
                    valid=True,
                    icon=item["icon"],
                    icontype="file"
                )
                log.debug("arg is: " + item["arg"])
                if "onenote:https" not in item["arg"]:
                    it.add_modifier('cmd',
                                    subtitle="open in OneNote",
                                    arg=item["url"] + ".one",
                                    valid=True)
                    it.setvar('theTitle', item["Name"])

            wf.send_feedback()
            return 0

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


def get_all_data(parent, prefix):
    """ recursively get every section of the parent

    this saves all the data in the global variable all_data()

    :param parent: entry point of search
    :param prefix: subtitle of page, pass as None for root
    :return: none

    """
    global subtitle
    global url
    global all_data
    global gosids

    if len(parent) > 0 and "Name" not in parent:
        for n in parent:
            get_all_data(n, prefix)
        prefix = None
        url = ""
        gosids = []

    if "Name" in parent:
        if "Children" in parent:
            gosids.append(parent["Gosid"])

            if prefix is None:
                pre = parent["Name"]
                url = "{0}{1}".format(urlbase, pre)
                subtitle = "{0}".format(parent["Name"])
                data = {
                    "Name": parent["Name"],
                    "subtitle": subtitle,
                    "arg": subtitle,
                    "autocomplete": parent["Name"],
                    "icon": "icons/notebook.png",
                    "url": url,
                    "gosids": gosids
                }
                all_data.append(data)

            else:
                pre = prefix + " > " + parent["Name"]
                subtitle = pre
                url = makeurl(subtitle)
                data = {
                    "Name": parent["Name"],
                    "subtitle": pre,
                    "arg": subtitle,
                    "autocomplete": parent["Name"],
                    "icon": "icons/section.png",
                    "url": url,
                    "gosids": gosids
                }
                all_data.append(data)

            get_all_data(parent["Children"], pre)

        else:
            subtitle = "{0} > {1}".format(prefix, parent["Name"])
            url = makeurl(subtitle)
            data = {
                "Name": parent["Name"],
                "subtitle": subtitle,
                "arg": url + ".one",
                "autocomplete": parent["Name"],
                "icon": "icons/page.png",
                "gosids": gosids
            }

            all_data.append(data)


def getAll(parent, prefix):
    """ recursively get every section of the parent

    :param parent: entry point of search
    :param prefix: subtitle of page, pass as None for root
    :return: none

    """
    global subtitle
    global url

    if len(parent) > 0 and "Name" not in parent:
        for n in parent:
            getAll(n, prefix)
        prefix = None
        url = ""

    if "Name" in parent:
        if "Children" in parent:

            if prefix is None:
                pre = parent["Name"]
                url = "{0}{1}".format(urlbase, pre)
                subtitle = "{0}".format(parent["Name"])

                it = wf.add_item(title=parent["Name"],
                                 subtitle=subtitle,
                                 arg=subtitle,
                                 autocomplete=parent["Name"],
                                 valid=True,
                                 icon="icons/notebook.png",
                                 icontype="file",
                                 quicklookurl=ONENOTE_ICON)
                it.add_modifier('cmd', subtitle="open in OneNote", arg=url, valid=True)
                it.setvar('theTitle', parent["Name"])

            else:
                pre = prefix + " > " + parent["Name"]
                subtitle = pre
                url = makeurl(subtitle)

                it = wf.add_item(title=parent["Name"],
                                 subtitle=pre,
                                 arg=subtitle,
                                 autocomplete=parent["Name"],
                                 valid=True,
                                 icon="icons/section.png",
                                 icontype="file",
                                 quicklookurl=ONENOTE_ICON)
                it.add_modifier('cmd', subtitle="open in OneNote", arg=url, valid=True)
                it.setvar('theTitle', parent["Name"])

            getAll(parent["Children"], pre)

        else:
            subtitle = "{0} > {1}".format(prefix, parent["Name"])
            url = makeurl(subtitle)
            it = wf.add_item(title=parent["Name"],
                             subtitle=subtitle,
                             arg=url + ".one",
                             autocomplete=parent["Name"],
                             valid=True,
                             icon="icons/page.png",
                             icontype="file",
                             quicklookurl=ONENOTE_ICON)


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
                         quicklookurl=ONENOTE_ICON)
        it.add_modifier('cmd', subtitle="open in OneNote", arg=url, valid=True)
        it.setvar('theTitle', n["Name"])
        it.setvar("q", sub)


def encode_url(url): return url.replace(" ", "%20")


def open_url(url):
    run_trigger('hide', wf.bundleid)    # hide alfred
    run_command(['open', encode_url(url)])


def clear_config():
    # unset_config('q')
    # unset_config('theTitle')
    log.info("'q' and 'theTitle' cleared")


if __name__ == "__main__":
    # db_test()
    wf = Workflow3(default_settings=DEFAULT_SETTINGS,
                   update_settings=UPDATE_SETTINGS,
                   help_url=HELP_URL)
    log = wf.logger
    sys.exit(wf.run(main))

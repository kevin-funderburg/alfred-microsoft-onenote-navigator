#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import argparse
import sqlite3
from sqlite3 import Error

import queries
from notebook_item import NotebookItem

from workflow import (Workflow3, ICON_INFO, ICON_WARNING,
                      ICON_ERROR, MATCH_ALL, MATCH_ALLCHARS,
                      MATCH_SUBSTRING, MATCH_STARTSWITH)
from workflow.util import run_trigger, run_command, set_config, unset_config
from workflow.background import run_in_background, is_running

__version__ = '2.0.0'

wf = None  # type: Workflow3
log = None

all_data = []
nitems =[]
path_map = {}

HELP_URL = 'https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator'
# GitHub repo for self-updating
UPDATE_SETTINGS = {'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator'}
DEFAULT_SETTINGS = {'urlbase': None}

ICON_ONENOTE = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"

ONENOTE_APP_SUPPORT = "~/Library/Containers/com.microsoft.onenote.mac/Data/Library/Application Support"
ONENOTE_FULL_SEARCH_PATH = ONENOTE_APP_SUPPORT + "/Microsoft User Data/OneNote/15.0/FullTextSearchIndex/"
ONENOTE_PLIST_PATH = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
ONENOTE_USER_UID = None
ONENOTE_PLIST = None
ALL_DB_PATHS = []

MERGED_DB = 'merged-onenote-data.db'

data = []


def get_page_path(ni):
    global path_map

    if ni.GOID not in path_map:
        if ni.Type == 4:
            entry = {ni.GOID: ni.Title}

        else:
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

            elif ni.Type < 3:
                parent_path = get_parent_path(ni)

            page_path = "{0}/{1}".format(parent_path, ni.Title)
            entry = {ni.GOID: page_path}

        path_map.update(entry)

    return path_map[ni.GOID]


def get_parent_row(parent_goid):
    conn = create_connection(wf.datafile(MERGED_DB))
    for row in conn.execute(queries.get_parent_row(parent_goid)):
        return row


def get_parent_path(item):
    global path_map
    if item.ParentGOID in path_map:
        return path_map[item.ParentGOID]
    else:
        parent_row = NotebookItem(get_parent_row(item.ParentGOID))
        return get_page_path(parent_row)


def get_row_by_guid(uid):
    res = run_query(queries.get_row_by_guid(uid))
    if res:
        return res[0]
    return None


def get_row_by_goid(goid):
    res = run_query(queries.get_row_by_goid(goid))
    if res:
        return res[0]
    return None


def get_children(ni):
    results = run_query(queries.get_children(ni))
    return results


def create_db():
    sql = queries.create_merged_db()
    conn = create_connection(wf.datafile(MERGED_DB))
    cur = conn.cursor()
    lines = sql.splitlines()
    for line in lines:
        if line is not "":
            cur.execute(str(line))

    conn.commit()


def update_db(): create_db()


def reset_db(): run_query(queries.reset_db())


def update_path_map():
    update_db()
    conn = create_connection(wf.datafile(MERGED_DB))
    for row in conn.execute(queries.get_all_items()):
        ni = NotebookItem(row)
        ni.set_path(get_page_path(ni))

    return path_map


def update_notebook_items():
    global nitems
    sql = "SELECT * FROM Entities;"
    log.debug(sql)
    results = run_query(sql)
    if not results:
        nitems = None
        # wf.add_item('No items', icon=ICON_WARNING)
    else:
        for result in results:
            item = NotebookItem(result)
            item.set_path(get_page_path(item))
            nitems.append(item)
            log.debug(item)

    wf.cache_data('notebook_items', nitems)
    return nitems


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
    log.debug("sql: {0}".format(sql))
    conn = create_connection(wf.datafile(MERGED_DB))
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


def get_results(sql):
    global nitems
    log.debug(sql)
    return run_query(sql)


def build_wf_items(results, args):
    if not results:
        wf.add_item('No items', icon=ICON_WARNING)
    else:
        for result in results:
            item = NotebookItem(result)
            item.set_path(get_page_path(item))
            nitems.append(item)
            uid = None  # set uid to None so alfred doesn't sort the results
            if args.all:
                uid = item.GUID
            it = wf.add_item(
                title=item.Title,
                arg=item.GUID,
                subtitle=item.path,
                uid=uid,
                autocomplete=item.Title,
                valid=True,
                icon=item.icon,
                icontype="file"
            )
            it.add_modifier('cmd', subtitle="open in OneNote", arg=item.url, valid=True)


def parse_args():
    parser = argparse.ArgumentParser(description="Search OneNote")
    parser.add_argument('-a', '--all', dest='all', action='store_true',
                        help='search unordered notebook items')
    parser.add_argument('-r', '--recent', dest='recent', action='store_true',
                        help='search recent notebook items')
    parser.add_argument('-m', '--modified', dest='modified', action='store_true',
                        help='search last modified notebook items')
    parser.add_argument('-n', '--notebooks', dest='notebooks', action='store_true',
                        help='search last modified notebook items')
    parser.add_argument('-u', '--update', dest='update', action='store_true',
                        help='update consolidated database')
    parser.add_argument('-o', '--open', dest='open', nargs='?', default=None,
                        help='open the notebook item with provided guid')
    parser.add_argument('-b', '--browse', dest='browse', action='store_true',
                        help='browse notebooks from top level')
    parser.add_argument('query', nargs='?', default=None)
    log.debug(wf.args)
    args = parser.parse_args(wf.args)
    return args


def encode_url(url):
    """encode url for shell to open

    :rtype: str
    """
    delims = {" ": "%20",
              "&": "%26",
              "{": "%7B",
              "}": "%7D"}
    for key, val in delims.items():
        url = url.replace(key, val)
    return url


def open_url(url):
    run_command(['open', encode_url(url)])


def clear_config():
    unset_config('q')
    log.info("'q' cleared")


def init_wf():
    global wf, log
    wf = Workflow3(default_settings=DEFAULT_SETTINGS,
                   update_settings=UPDATE_SETTINGS,
                   help_url=HELP_URL)
    log = wf.logger


def main(wf):
    global path_map
    log.debug('Started workflow')
    log.debug("os.getenv('q'): {0}".format(os.getenv('q')))
    args = parse_args()

    if wf.update_available:
        # Add a notification to top of Script Filter results
        wf.add_item('New version available',
                    'Action this item to install the update',
                    autocomplete='workflow:update',
                    icon=ICON_INFO)

    path_map = wf.cached_data('path_map', update_path_map, max_age=300)
    if not wf.cached_data_fresh('path_map', 300):
        log.debug('data is stale, updating DB')
        update_db()
    # results = wf.cached_data('notebook_items', update_notebook_items, max_age=600)

    if args.browse:
        log.debug("arg type: args.browse")
        goid = os.getenv('q')
        ni = NotebookItem(get_row_by_goid(goid))
        log.debug("ni.GOID: {0}".format(ni.GOID))
        results = get_children(ni)
        build_wf_items(results, args)
        wf.send_feedback()
        return 0

    if args.open:
        log.debug("arg type: args.open")
        ni = NotebookItem(get_row_by_guid(args.open))

        if ni.Type == 1:
            log.debug('argument is a page, preparing to open...' + ni.Title)
            run_trigger('hide')
            open_url(ni.url)
        else:
            log.info("setting 'q' to: {0}".format(args.open))
            log.debug('argument is not a page, gathering sub-pages...')
            set_config("q", ni.GOID)
            run_trigger('1b')
        return 0

    if args.update:
        update_db()
        wf.add_item('Updating OneNote cache.',
                    'Give me a minute and then try searching again.',
                    valid=False,
                    icon=ICON_INFO)
        wf.send_feedback()
        return 0

    query = populate_query(args)
    results = get_results(query)
    build_wf_items(results, args)
    wf.send_feedback()


if __name__ == "__main__":
    init_wf()
    sys.exit(wf.run(main))

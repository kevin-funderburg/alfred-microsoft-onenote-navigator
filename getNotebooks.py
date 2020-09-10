#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
import argparse
import plistlib

from workflow import Workflow3, ICON_INFO, ICON_WARNING, \
    ICON_ERROR, MATCH_ALL, MATCH_ALLCHARS, MATCH_SUBSTRING
from workflow.util import run_trigger, run_applescript, run_command

__version__ = '1.3.1'

wf = None
log = None
sub = []
subtitle = ""
url = ""
urlbase = None
all_data = []

HELP_URL = 'https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator'

# GitHub repo for self-updating
UPDATE_SETTINGS = {'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator'}

DEFAULT_SETTINGS = {'urlbase': "null"}

ONENOTE_PLIST = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist")
ICON_APP = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"
DATA_DIR = "~/Library/Application Support/Alfred/Workflow Data/com.kfunderburg.oneNoteNav/"

try:
    os.mkdir(os.path.expanduser(DATA_DIR))
except OSError:
    pass

DATA_FILE = os.path.expanduser(DATA_DIR) + "data.plist"
SETTINGS_PATH = os.path.expanduser(DATA_DIR) + "settings.plist"


def main(wf):
    global urlbase

    if not os.path.exists(SETTINGS_PATH):
        log.info("no settings set, creating default settings")
        plistlib.writePlist(DEFAULT_SETTINGS, SETTINGS_PATH)

    settings = plistlib.readPlist(SETTINGS_PATH)

    if wf.update_available:
        # Add a notification to top of Script Filter results
        wf.add_item('New version available',
                    'Action this item to install the update',
                    autocomplete='workflow:update',
                    icon=ICON_INFO)
    # build argument parser to parse script args and collect their
    # values
    parser = argparse.ArgumentParser()
    # --seturl argument and save its value to 'urlbase' (dest).
    parser.add_argument('--seturl', dest='urlbase', nargs='?', default=None)
    parser.add_argument('--browse', dest='browse', nargs='?', default=None)
    parser.add_argument('--write', dest='write', nargs='?', default=None)
    parser.add_argument('--type', dest='type', nargs='?', default=None)
    parser.add_argument('--warn', dest='warn', nargs='?', default=None)
    parser.add_argument('--url', dest='url', nargs='?', default=None)
    parser.add_argument('query', nargs='?', default=None)
    args = parser.parse_args(wf.args)


    ####################################################################
    # Save the provided URL base
    ####################################################################
    if args.urlbase:  # Script was passed as a URL
        set_url(args.urlbase)
        return 0

    ####################################################################
    # Warn if bad url was passed in
    ####################################################################
    if args.warn:
        warn(args.warn)
        return 0

    if args.url:
        if 'onenote:https' in args.url:
            open_url(args.url)
        else:
            log.info("setting 'q' to: {0}".format(args.url))
            wf.setvar("q", args.url, True)
            aps = "tell application id \"com.runningwithcrayons.Alfred\" to " \
                  "run trigger \"browse_section\" in workflow \"com.kfunderburg.oneNoteNav\""
            run_applescript(aps)
            # run_trigger('browse_section', wf.bundleid)
        return 0

    ####################################################################
    # Check that we have a URL saved
    ##################################################################
    if settings['urlbase'] == "null":
        wf.add_item('No URL set yet.',
                    'Please use seturl to set your OneNote url base.',
                    valid=False,
                    icon=ICON_WARNING)
        wf.send_feedback()
        return 0

    # get plist data from OneNote plist
    onenote_pl = plistlib.readPlist(ONENOTE_PLIST)
    urlbase = settings['urlbase']

    if args.type == 'searchall':
        get_all_data(onenote_pl, None)

        items = wf.filter(args.query, all_data, key_for_data, match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING)

        if not items:
            wf.add_item(title='No matches',
                        icon=ICON_WARNING)

        for item in items:
            it = wf.add_item(title=item["Name"],
                             subtitle=item["subtitle"],
                             arg=item["arg"],
                             autocomplete=item["Name"],
                             valid=item["valid"],
                             icon=item["icon"],
                             icontype=item["icon"],
                             quicklookurl=ICON_APP)
            if "onenote:https" not in item["arg"]:
                it.add_modifier('cmd', subtitle="open in OneNote", arg=item["url"] + ".one", valid=True)
                it.setvar('theTitle', item["Name"])

        wf.send_feedback()
        return 0

    elif args.browse:
        log.debug("browse is: {0}".format(args.browse))
        log.debug("query is: {0}".format(args.query))

        if args.browse == 'notebooks':
            data = get_notebook_data()
            items = wf.filter(args.query, data, key_for_data, match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING)

            if not items:
                wf.add_item(title='No matches',
                            icon=ICON_WARNING)

            for item in items:
                it = wf.add_item(title=item["Name"],
                                 subtitle=item["subtitle"],
                                 arg="",
                                 autocomplete=item["Name"],
                                 valid=True,
                                 icon=item["icon"],
                                 icontype="file",
                                 quicklookurl=ICON_APP)
                it.add_modifier('cmd', subtitle="open in OneNote", arg=item["arg"], valid=True)
                it.setvar('theTitle', item["Name"])
                it.setvar("q", item["subtitle"])

            wf.send_feedback()
            return 0

        else:
            data = browse_child()
            items = wf.filter(args.query, data, key_for_data, match_on=MATCH_ALL ^ MATCH_ALLCHARS ^ MATCH_SUBSTRING)

            if not items:
                wf.add_item(title='No matches',
                            icon=ICON_WARNING)

            for item in items:
                it = wf.add_item(title=item["Name"],
                                 subtitle=item["subtitle"],
                                 arg=item["arg"],
                                 autocomplete=item["Name"],
                                 valid=True,
                                 icon=item["icon"],
                                 icontype="file",
                                 quicklookurl=ICON_APP)
                log.debug("arg is: " + item["arg"])
                if "onenote:https" not in item["arg"]:
                    it.add_modifier('cmd', subtitle="open in OneNote", arg=item["url"] + ".one", valid=True)
                    it.setvar('theTitle', item["Name"])

            wf.send_feedback()
            return 0

    return 0


def key_for_data(data):
    return '{}'.format(data['Name'])


def get_notebook_data():
    onenote_pl = plistlib.readPlist(ONENOTE_PLIST)
    data = []
    for n in onenote_pl:
        sub = "{0}".format(n["Name"])
        log.debug("sub is: " + sub)
        url = makeurl(sub)   # makeurl(sub)
        item = {
            "Name": n["Name"],
            "subtitle": sub,
            "arg": url,
            "autocomplete": n["Name"],
            "valid": True,
            "icon": "icons/notebook.png",
            "icontype": "file",
            "quicklookurl": ICON_APP
        }
        data.append(item)
    return data


def get_all_data(parent, prefix):
    """ recursively get every section of the parent

    :param parent: entry point of search
    :param prefix: subtitle of page, pass as None for root
    :return: none

    """
    global subtitle
    global url
    global all_data

    if len(parent) > 0 and "Name" not in parent:
        for n in parent:
            get_all_data(n, prefix)
        prefix = None
        url = ""

    if "Name" in parent:
        if "Children" in parent:

            if prefix is None:
                pre = parent["Name"]
                url = "{0}{1}".format(urlbase, pre)
                subtitle = "{0}".format(parent["Name"])
                data = {
                    "Name": parent["Name"],
                    "subtitle": subtitle,
                    "arg": subtitle,
                    "autocomplete": parent["Name"],
                    "valid": True,
                    "icon": "icons/notebook.png",
                    "icontype": "file",
                    "quicklookurl": ICON_APP,
                    "url": url
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
                    "valid": True,
                    "icon": "icons/section.png",
                    "icontype": "file",
                    "quicklookurl": ICON_APP,
                    "url": url
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
                "valid": True,
                "icon": "icons/page.png",
                "icontype": "file",
                "quicklookurl": ICON_APP
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
                                 quicklookurl=ICON_APP)
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
                                 quicklookurl=ICON_APP)
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
                             quicklookurl=ICON_APP)


def makeurl(prefix):
    newurl = "{0}{1}".format(urlbase, prefix.replace(" > ", "/"))
    newurl = newurl.replace(" ", "%20")
    return newurl


def get_child(childstr):
    """ finds the child of the onenote plist

    :param childstr: path of child, formatted as '[element] > [element]'
    :return: child element of ooenote plist
    """
    items = childstr.split(" > ")
    pl = plistlib.readPlist(ONENOTE_PLIST)
    child = None
    for x in range(len(items)):
        if x == 0:
            for p in pl:
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
    :return: none
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
                "valid": True,
                "icon": "icons/section.png",
                "icontype": "file",
                "quicklookurl": ICON_APP,
                "url": url
            }

        else:
            item = {
                "Name": c["Name"],
                "subtitle": sub,
                "arg": url + ".one",
                "autocomplete": c["Name"],
                "valid": True,
                "icon": "icons/page.png",
                "icontype": "file",
                "quicklookurl": ICON_APP
            }

        data.append(item)

    return data


def set_url(urlbase):
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
        plistlib.writePlist({"urlbase": urlbase}, SETTINGS_PATH)
        # notify("OneNote URL setup", "OneNote URL base was stored successfully")
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
    onenote_pl = plistlib.readPlist(ONENOTE_PLIST)
    for n in onenote_pl:
        sub = "{0}".format(n["Name"])
        url = makeurl(sub)
        it = wf.add_item(title=n["Name"],
                         subtitle=sub,
                         arg="",
                         autocomplete=n["Name"],
                         valid=True,
                         icon="icons/notebook.png",
                         icontype="file",
                         quicklookurl=ICON_APP)
        it.add_modifier('cmd', subtitle="open in OneNote", arg=url, valid=True)
        it.setvar('theTitle', n["Name"])
        it.setvar("q", sub)


def open_url(url):
    # aps = "tell application id \"com.runningwithcrayons.Alfred\" to" \
    #       " run trigger \"hide\" in workflow \"com.kfunderburg.oneNoteNav\"" \
    #       "with argument \"test\""
    # run_applescript(aps)
    url.replace(" ", "%20")
    run_trigger('hide', wf.bundleid)  # hide alfred
    run_command(['open', url])
    aps = "tell application id \"com.runningwithcrayons.Alfred\" to " \
          "remove configuration \"q\" in workflow \"com.kfunderburg.oneNoteNav\""
    run_applescript(aps)
    log.info("q var cleared")


if __name__ == "__main__":
    wf = Workflow3(help_url=HELP_URL)
    log = wf.logger
    sys.exit(wf.run(main))

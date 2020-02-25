#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
import argparse
import plistlib

from workflow import Workflow3, ICON_INFO, ICON_WARNING, ICON_ERROR

__version__ = '1.3.0'

wf = None
log = None
sub = []
subtitle = ""
url = ""
urlbase = None
# prefix = None

HELP_URL = 'https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator'

# GitHub repo for self-updating
UPDATE_SETTINGS={'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator'}

DEFAULT_SETTINGS = {'urlbase': "null"}

ONENOTE_PLIST = os.path.expanduser("~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist")
ICON_APP = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"
DATA_DIR = "~/Library/Application Support/Alfred/Workflow Data/com.kfunderburg.oneNoteNav/"

try:
    os.mkdir(os.path.expanduser(DATA_DIR))
except OSError:
    pass

DATA_FILE = os.path.expanduser(DATA_DIR) + "data.plist"
SETTINGS_PATH = os.path.expanduser(DATA_DIR) + "settings.plist"


def main(wf):
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
    parser.add_argument('--write', dest='write', nargs='?', default=None)
    parser.add_argument('--type', dest='type', nargs='?', default=None)
    parser.add_argument('--warn', dest='warn', nargs='?', default=None)
    parser.add_argument('query', nargs='?', default=None)
    args = parser.parse_args(wf.args)

    ####################################################################
    # Save the provided URL base
    ####################################################################
    if args.urlbase:  # Script was passed as a URL
        log.debug("arg is url: {0}".format(args.urlbase))
        if 'sharepoint' in args.urlbase:
            # sharepoint URLs will not work, pass as an error
            print("sharepoint")
            return 0
        if 'onenote:https' in args.urlbase:
            # extract onenote url
            args.urlbase = re.search('(onenote:.*)', args.urlbase).group(1)
            url = args.urlbase.split("/")
            urlbase = url[0] + "//" + url[2] + "/" + url[3] + "/" + url[4] + "/"
            # write to plist
            plistlib.writePlist({"urlbase": urlbase}, SETTINGS_PATH)
            # tell alfred how to display a success notification
            print("true")
        else:
            # tell alfred how to display alert that the url is bad
            print("invalid")
        return 0

    ####################################################################
    # Warn if bad url was passed in
    ####################################################################
    if args.warn:
        if "sharepoint" in args.warn:
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
    # plist = plistlib.writePlistToString(onenote_pl)
    # log.info(plist)
    global urlbase
    urlbase = settings['urlbase']

    if args.type == 'searchall':
        getAll(onenote_pl, None)
        # children = os.getenv('children')
        # if children is None:
        #     getAll(onenote_pl, None)
        # else:
        #     children = plistlib.readPlistFromString(children)
        #     getAll(children, None)
        wf.send_feedback()
        return 0

    if args.type == 'browse':
        log.debug("arg is write: {0}".format(args.type))

        browse_child()
        wf.send_feedback()

        return 0

    if args.type == 'notebook':
        # write all notebook plist data to data.plist for later iterations
        plistlib.writePlist(onenote_pl, DATA_FILE)

        for n in onenote_pl:
            it = wf.add_item(title=n["Name"],
                             subtitle="View " + n["Name"] + "'s sections",
                             arg=n["Name"],
                             autocomplete=n["Name"],
                             valid=True,
                             icon="icon.png",
                             icontype="file",
                             quicklookurl=ICON_APP)
            it.setvar('subPreFix', "")
            it.setvar('notebook', n["Name"])
            it.setvar('theURL', "")
            it.add_modifier('cmd',
                            subtitle="{0}{1}".format(settings['urlbase'], n["Name"]),
                            arg="{0}{1}".format(settings['urlbase'], n["Name"]),
                            valid=True)

        if len(wf._items) == 0:
            wf.add_item('No notebooks found', icon=ICON_WARNING)
            wf.send_feedback()
            return 0

        wf.send_feedback()
        return 0

    if args.type == 'section':
        notebook = os.getenv('notebook')
        q = os.getenv('q')

        if notebook is None:
            notebook = q

        found = False
        pl = plistlib.readPlist(DATA_FILE)
        for p in pl:
            if p["Name"] == q:
                found = True
                break

        if not found:
            wf.add_item('No section ' + q + 'was found', icon=ICON_WARNING)
            wf.send_feedback()
            return 0

        if "Children" in p:
            # write children of current section to data.plist for further iterations
            plistlib.writePlist(p["Children"], DATA_FILE)
            for c in p["Children"]:
                subPreFix = os.getenv("subPreFix")
                theURL = os.getenv("theURL")

                if (subPreFix == "") or (subPreFix is None):
                    subPreFix = "{0} > {1}".format(notebook, c["Name"])
                else:
                    subPreFix = "{0} > {1}".format(subPreFix, c["Name"])

                if (theURL == "") or (theURL is None):
                    theURL = "{0}{1}/{2}".format(settings['urlbase'], notebook, c["Name"])
                else:
                    theURL = "{0}/{1}".format(theURL, c["Name"])

                it = wf.add_item(title=c["Name"],
                                 subtitle=subPreFix,
                                 arg=c["Name"],
                                 autocomplete=c["Name"],
                                 valid=True,
                                 icon="icons/section.png",
                                 icontype="file",
                                 quicklookurl=ICON_APP)
                it.setvar('subPreFix', subPreFix)
                it.setvar('theURL', theURL)
                it.setvar('leaf', "false")
                it.add_modifier('cmd', subtitle=theURL, arg=theURL, valid=True)
        else:
            # section has no children so can't go any deeper
            theURL = os.getenv("theURL")
            theURL += ".one"

            it = wf.add_item(title=p["Name"],
                             subtitle=theURL,
                             arg=theURL,
                             autocomplete=p["Name"],
                             valid=True,
                             icon="icons/page.png",
                             icontype="file",
                             quicklookurl=ICON_APP)
            it.setvar('leaf', "true")

        if len(wf._items) == 0:
            wf.add_item('No sections found', icon=ICON_WARNING)
            wf.send_feedback()
            return 0

        wf.send_feedback()
        return 0

    wf.add_item('No workflow items were made', icon=ICON_WARNING)
    wf.send_feedback()
    return 0


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
                # it.setvar('children', plistlib.writePlistToString(parent["Children"]))

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
                # it.setvar('children', plistlib.writePlistToString(parent["Children"]))

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
            # it.setvar('child', subtitle)
            # it.add_modifier('cmd', subtitle="browse in Alfred", arg=subtitle, valid=True)


def makeurl(prefix):
    newurl = "{0}{1}".format(urlbase, prefix.replace(" > ", "/"))
    newurl = newurl.replace(" ", "%20")
    return newurl


def get_child(childstr):
    itms = childstr.split(" > ")
    onenote_pl = plistlib.readPlist(ONENOTE_PLIST)

    pl = onenote_pl
    child = None

    for x in range(len(itms)):
        if x == 0:
            for p in pl:
                if p["Name"] == itms[0]:
                    log.info("found")
                    notebook = p
                    child = notebook["Children"]
                    break
        else:
            for c in child:
                if c["Name"] == itms[x]:
                    if "Children" in c:
                        child = c["Children"]
                    break

    log.info("child found")
    return child


def browse_child():
    q = os.getenv('q')
    # q = "Work > Career"
    child = get_child(q)
    for c in child:
        sub = "{0} > {1}".format(q, c["Name"])
        url = makeurl(sub)
        if "Children" in c:
            it = wf.add_item(title=c["Name"],
                             subtitle=sub,
                             arg=sub,
                             autocomplete=c["Name"],
                             valid=True,
                             icon="icons/section.png",
                             icontype="file",
                             quicklookurl=ICON_APP)
            it.add_modifier('cmd', subtitle="open in OneNote", arg=url + ".one", valid=True)

        else:
            it = wf.add_item(title=c["Name"],
                             subtitle=sub,
                             arg=url + ".one",
                             autocomplete=c["Name"],
                             valid=True,
                             icon="icons/page.png",
                             icontype="file",
                             quicklookurl=ICON_APP)


if __name__ == "__main__":
    wf = Workflow3(help_url=HELP_URL)
    log = wf.logger
    sys.exit(wf.run(main))

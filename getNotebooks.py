#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
import argparse
from workflow import Workflow3, ICON_INFO, ICON_WARNING

__version__ = '1.1'

wf = None
log = None
HELP_URL = ('https://github.com/kevin-funderburg/alfred-microsoft-onenote-navigator',
            '#alfred-microsft-onenote-navigator')

UPDATE_SETTINGS={
        'github_slug': 'kevin-funderburg/alfred-microsoft-onenote-navigator',
}
# DEFAULT_DATA_FILE=
ONENOTEPLIST = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
ICON_APP = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"


def main(wf):

    import plistlib

    if wf.update_available:
        # Add a notification to top of Script Filter results
        wf.add_item('New version available',
                    'Action this item to install the update',
                    autocomplete='workflow:update',
                    icon=ICON_INFO)

    # build argument parser to parse script args and collect their
    # values
    parser = argparse.ArgumentParser()
    # add an optional (nargs='?') --seturl argument and save its
    # value to 'urlbase' (dest). This will be called from a separate "Run Script"
    # action with the API key
    parser.add_argument('--seturl', dest='urlbase', nargs='?', default=None)
    # parser.add_argument('--notebooks', dest='notebooks', nargs='?', default=None)
    parser.add_argument('--type', dest='type', nargs='?', default=None)
    # parser.add_argument('--sections', dest='sections', nargs='?', default=None)
    # add an optional query and save it to 'query'
    parser.add_argument('query', nargs='?', default=None)
    # parse the script's arguments
    args = parser.parse_args(wf.args)

    ####################################################################
    # Save the provided URL base
    ####################################################################

    # decide what to do based on arguments
    if args.urlbase:  # Script was passed as a URL
        if 'onenote:https' in args.urlbase:
            args.urlbase = re.search('(onenote.*/Documents/).*', args.urlbase).group(1)
            # save the URL
            wf.settings['urlbase'] = args.urlbase
            return 0
        else:
            wf.add_item("The argument is not a OneNote URL string",
                        "right click a OneNote page and click 'Copy Link to Page'",
                        valid=False,
                        icon=ICON_WARNING)
            wf.send_feedback()
            return 0

    ####################################################################
    # Check that we have a URL saved
    ####################################################################
    try:
        urlbase = wf.settings.get('urlbase', None)
    except "URL Not Found":
            wf.add_item('No URL set yet.',
                        'Please use urlsetkey to set your OneNote url base.',
                        valid=False,
                        icon=ICON_WARNING)
            wf.send_feedback()
            return 0

    if args.type == 'notebook':
        one_note_pl = plistlib.readPlist(os.path.expanduser(ONENOTEPLIST))
        # write all notebook plist data to data.plist for getNotebookSections.py
        plistlib.writePlist(one_note_pl, wf.datafile('data.plist'))

        for n in one_note_pl:
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
            it.add_modifier('cmd', subtitle=urlbase + n["Name"], arg=urlbase + n["Name"], valid=True)

        # if len(wf.items) == 0:
        #     wf.add_item('No notebooks found', icon=ICON_WARNING)
        #     wf.send_feedback()
        #     return 0

        wf.send_feedback()
        return 0

    if args.type == 'section':
        notebook = os.getenv('notebook')
        q = os.getenv('q')
        if notebook is None:
            notebook = q

        found = False
        for p in pl:
            if p["Name"] == q:
                found = True
                break

        if not found:
            log.exception('section was not found')
            pass

        if "Children" in p:
            # write children of current section to data.plist for further iterations
            plistlib.writePlist(p["Children"], wf.datafile('data.plist'))
            for c in p["Children"]:
                subPreFix = os.getenv("subPreFix")
                theURL = os.getenv("theURL")

                if (subPreFix == "") or (subPreFix is None):
                    subPreFix = notebook + " > " + c["Name"]
                else:
                    subPreFix = subPreFix + " > " + c["Name"]

                if (theURL == "") or (theURL is None):
                    theURL = urlbase + notebook + "/" + c["Name"]
                else:
                    theURL = theURL + "/" + c["Name"]

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

        if len(wf.items) == 0:
            wf.add_item('No sections found', icon=ICON_WARNING)
            wf.send_feedback()
            return 0

        wf.send_feedback()
        return 0

    wf.add_item('No workflow items were made', icon=ICON_WARNING)
    wf.send_feedback()
    return 0






if __name__ == "__main__":
    # wf = Workflow3(update_settings=UPDATE_SETTINGS,
    #                help_url=HELP_URL)
    wf = Workflow3(help_url=HELP_URL)
    log = wf.logger
    sys.exit(wf.run(main))

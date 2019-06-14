#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
from workflow import Workflow3

wf = None
log = None

# NOTE:
# URLBASE is unique to each user, so this needs to be updated to your system
# How to update BASEURL:
# 1. Right-click a notebook in OneNote and choose "Copy link to Notebook"
# That will copy 2 links, a web version and a mac version, so to get the mac
# version,
# 2. Paste into a some text file then
# 3. Copy from the beginning of the url that begins with 'onenote:https:'
# through '/Documents/' and replace the URLBASE with what you copied.
URLBASE = "onenote:https://d.docs.live.net/9478a1a4ec3795b7/Documents/"
APPICON = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"
ONENOTEPLISTPATH = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"

def main(wf):

    import plistlib

    log.info(wf.datadir)
    log.info('Workflow response complete')

    dataPlist = wf.datadir + "/data.plist"
    pl = plistlib.readPlist(os.path.expanduser(dataPlist))

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
        plistlib.writePlist(p["Children"], dataPlist)
        for c in p["Children"]:
            subPreFix = os.getenv("subPreFix")
            theURL = os.getenv("theURL")

            if (subPreFix == "") or (subPreFix is None):
                subPreFix = notebook + " > " + c["Name"]
            else:
                subPreFix = subPreFix + " > " + c["Name"]

            if (theURL == "") or (theURL is None):
                theURL = URLBASE + notebook + "/" + c["Name"]
            else:
                theURL = theURL + "/" + c["Name"]

            # # it = wf.add_item(title=c["Name"],
            #                  subtitle=subPreFix,
            #                  arg=c["Name"],
            #                  autocomplete=c["Name"],
            #                  valid=True,
            #                  icon=APPICON,
            #                  icontype="file",
            #                  quicklookurl=APPICON)
            it = wf.add_item(title=c["Name"],
                             subtitle=subPreFix,
                             arg=c["Name"],
                             autocomplete=c["Name"],
                             valid=True,
                             icon="icons/section.png",
                             icontype="file",
                             quicklookurl=APPICON)
            it.setvar('subPreFix', subPreFix)
            it.setvar('theURL', theURL)
            it.setvar('leaf', "false")
            it.add_modifier('cmd', subtitle=theURL, arg=theURL, valid=True)
    else:
        # section has no children so can't go any deeper
        theURL = os.getenv("theURL")
        theURL += ".one"

        # it = wf.add_item(title=p["Name"],
        #                  subtitle=theURL,
        #                  arg=theURL,
        #                  autocomplete=p["Name"],
        #                  valid=True,
        #                  icon=APPICON,
        #                  icontype="file",
        #                  quicklookurl=APPICON)
        it = wf.add_item(title=p["Name"],
                         subtitle=theURL,
                         arg=theURL,
                         autocomplete=p["Name"],
                         valid=True,
                         icon="icons/page.png",
                         icontype="file",
                         quicklookurl=APPICON)
        it.setvar('leaf', "true")

    return wf.send_feedback()


if __name__ == "__main__":
    wf = Workflow3()
    log = wf.logger
    sys.exit(wf.run(main))

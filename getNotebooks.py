#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
from workflow import Workflow3, ICON_INFO

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
ONENOTEPLIST = "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"
APPICON = "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"

def main(wf):

    import plistlib

    log.info(wf.datadir)
    log.info('Workflow response complete')

    dataPlist = wf.datadir + "/data.plist"

    if not os.path.exists(dataPlist):
        log.info('data.plist missing, creating empty plist')
        data = {}
        plistlib.writePlist(data, dataPlist)

    if "URLbase" not in dataPlist:
        it = wf.add_item(title="OneNote URL not found",
                         subtitle="Copy a link from a notebook and paste it into the helper",
                         icon=ICON_INFO)
    else:
        pl = plistlib.readPlist(os.path.expanduser(ONENOTEPLIST))
        data = {'URLbase': URLBASE,
                'notebooks': pl}
        # pl2 = pl
        # pl2["URLbase"] = URLBASE
        plistlib.writePlist(data, dataPlist)  # write all notebook plist data to data.plist for getNotebookSections.py

        for n in data["notebooks"]:
            it = wf.add_item(title=n["Name"],
                             subtitle="View " + n["Name"] + "'s sections",
                             arg=n["Name"],
                             autocomplete=n["Name"],
                             valid=True,
                             icon="icon.png",
                             icontype="file",
                             quicklookurl=APPICON)
            it.setvar('subPreFix', "")
            it.setvar('notebook', n["Name"])
            it.setvar('theURL', "")
            it.add_modifier('cmd', subtitle=URLBASE + n["Name"], arg=URLBASE + n["Name"], valid=True)

    return wf.send_feedback()


# def config(wf):



if __name__ == "__main__":
    wf = Workflow3()
    log = wf.logger
    sys.exit(wf.run(main))

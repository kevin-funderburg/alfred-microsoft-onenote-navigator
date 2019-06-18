#!/usr/bin/python
# encoding: utf-8

from __future__ import unicode_literals

import os
import sys
import re
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

    log.info('Creating URLbase')

    # build argument parser to parse script args and collect their
    # values
    parser = argparse.ArgumentParser()
    # add an optional (nargs='?') --seturl argument and save its
    # value to 'urlbase' (dest). This will be called from a separate "Run Script"
    # action with the API key
    parser.add_argument('--setkey', dest='urlbase', nargs='?', default=None)
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

    urlbase = wf.settings.get('urlbase', None)
    if not urlbase:
        wf.add_item('No URL set yet.',
                    'Please use urlsetkey to set your OneNote url base.',
                    valid=False,
                    icon=ICON_WARNING)
        wf.send_feedback()
        return 0



    argURL = wf.args[0]
    log.debug('args: ' + argURL)

    if argURL.startswith('onenote'):
        URLbase = re.search('(onenote.*/Documents/).*', argURL).group(1)

    wf.settings['URLbase'] = URLbase



    ####################################################################
    # Check that we have an API key saved
    ####################################################################

    urlbase = wf.settings.get('URLbase', None)
    if not urlbase:
        wf.add_item('No URL set yet.',
                    'Please use pbsetkey to set your Pinboard API key.',
                    valid=False,
                    icon=ICON_WARNING)
        wf.send_feedback()
        return 0

    pl = plistlib.readPlist(os.path.expanduser(ONENOTEPLIST))
    plistlib.writePlist(data, wf.datafile('data.plist'))

    it = wf.add_item(title='Done!',
                     subtitle="You can now open your notebooks/sections/pages")

    return wf.send_feedback()


if __name__ == "__main__":
    wf = Workflow3()
    log = wf.logger
    sys.exit(wf.run(main))

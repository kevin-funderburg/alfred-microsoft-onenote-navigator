use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions
use framework "Foundation"

property urlbase : "onenote:https://d.docs.live.net/9478a1a4ec3795b7/Documents/"
property extensionPlist : POSIX path of (path to library folder from user domain) & "/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"

on run argv
	
	if class of argv = script then -- This condition is for troubleshooting within Script Debugger
		set wfPath to "/Users/kevinfunderburg/Dropbox/Library/Application Support/Alfred/Alfred.alfredpreferences/workflows/user.workflow.F656F39C-B1D5-471E-942E-8D76BDDBE40A"
		
		# load the Workflow library
		set wlib to load script "/Users/kevinfunderburg/Dropbox/Library/Application Support/Alfred/Alfred.alfredpreferences/workflows/user.workflow.F656F39C-B1D5-471E-942E-8D76BDDBE40A/q_workflow.scpt" as POSIX file
		
		# create a new Workflow Class
		set wf to wlib's new_workflow()
		
	else
		# get the workflow's source folder
		set workflowFolder to do shell script "pwd"
		
		# load the Workflow library
		set wlib to load script workflowFolder & "/q_workflow.scpt" as POSIX file
		
		# create a new Workflow Class
		set wf to wlib's new_workflow()
		
		tell wf
			set uid to wf's getUID()
			set alfredPrefs to wf's getPreferences()
			set wfPath to alfredPrefs & "/workflows/" & uid
		end tell
	end if
	
	set dataPlist to wfPath & "/data.plist"
	
	tell application "System Events"
		tell property list file extensionPlist
			set theRecord to value of every property list item
		end tell
		tell property list file dataPlist
			tell contents
				set value to theRecord
			end tell
		end tell
	end tell
	
	--set theNames to my valueForKey:"name" inList:theRecord
	repeat with n in items of theRecord
		if |name| of n is "Mac" then
			set iconPath to "icons/iMac.png"
		else if |name| of n is "School" then
			set iconPath to "icons/lecture.png"
		else if |name| of n is "Kevin's Notebook" then
			set iconPath to "icons/switchUser@2x.png"
		else if |name| of n is "Cooking" then
			set iconPath to "icons/cooking.png"
		else if |name| of n is "Work" then
			set iconPath to "icons/work.png"
		else
			set iconPath to "/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns"
		end if
		
		add_result of wf given theUid:(|name| of n), theArg:|name| of n, theTitle:|name| of n, theSubtitle:"View " & |name| of n & " notebook sections", theIcon:{theType:missing value, thePath:iconPath}, theAutocomplete:|name| of n, theType:missing value, isValid:"yes", theQuicklook:"/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns", theVars:{{|name|:"subPreFix", value:""}, {|name|:"notebook", value:|name| of n}, {|name|:"theURL", value:""}}, theMods:{cmd:{isValid:"yes", theArg:urlbase & |name| of n, theSubtitle:urlbase & |name| of n}, alt:missing value}
	end repeat
	
	return wf's to_json("")
end run
use AppleScript version "2.4" -- Yosemite (10.10) or later
use scripting additions

property urlbase : "onenote:https://d.docs.live.net/9478a1a4ec3795b7/Documents/"
property notebooksplist : "~/Library/Group Containers/UBF8T346G9.Office/OneNote/ShareExtension/Notebooks.plist"

---——————————————————————————————————————————————-
on run argv
	--—————————————————————————————————————————————-	
	if class of argv = script then -- This if block is to troubleshoot within Script Debugger, not Alfred
		set wfPath to "/Users/kevinfunderburg/Dropbox/Library/Application Support/Alfred/Alfred.alfredpreferences/workflows/user.workflow.F656F39C-B1D5-471E-942E-8D76BDDBE40A"
		
		# load the Workflow library
		set wlib to load script "Macintosh HD:Users:kevinfunderburg:Dropbox:Library:Application Support:Alfred:Alfred.alfredpreferences:workflows:user.workflow.F656F39C-B1D5-471E-942E-8D76BDDBE40A:q_workflow.scpt" as alias
		
		# create a new Workflow Class
		set wf to wlib's new_workflow()
		
	else
		# get the workflow's source folder
		set workflowFolder to do shell script "pwd"
		
		# load the Workflow library
		set wlib to load script workflowFolder & "/q_workflow.scpt" as POSIX file
		
		# create a new Workflow Class
		set wf to wlib's newWorkflow()
		
		tell wf
			set uid to wf's getUID()
			set alfredPrefs to wf's getPreferences()
			set notebook to wf's getVar("notebook")
			set q to wf's getVar("q")
			set wfPath to alfredPrefs & "/workflows/" & uid
		end tell
	end if
	
	set dataPlist to wfPath & "/data.plist"
	
	log "theArg is: " & q
	
	if notebook is missing value then set notebook to q
	
	tell application "System Events"
		tell property list file dataPlist
			set theRecord to value of every property list item
		end tell
	end tell
	
	--» Get the record of the chosen notebook
	set found to false
	repeat with n in theRecord
		if |name| of n is q then
			set found to true
			exit repeat
		end if
	end repeat
	if found is false then error "section not found"
	
	try
		tell application "System Events"
			tell property list file dataPlist
				tell contents
					set value to Children of n
				end tell
			end tell
		end tell
		
		repeat with c in (Children of n)
			--log "CHILD IS: " & (|name| of c)
			set subPreFix to wf's getVar("subPreFix")
			set theURL to wf's getVar("theURL")
			--log "SUBPREFIX IS: " & subPreFix
			if subPreFix is "" or subPreFix is missing value then
				set subPreFix to notebook & " > " & |name| of c
			else
				set subPreFix to subPreFix & " > " & |name| of c
			end if
			if theURL is "" or theURL is missing value then
				set theURL to urlbase & notebook & "/" & |name| of c
			else
				set theURL to theURL & "/" & |name| of c
			end if
			
			if theURL contains "AppleScript" then
				set iconPath to "icons/appleScript.png"
			else
				set iconPath to "icons/section.png"
			end if
			
			add_result of wf given theUid:(|name| of c), theArg:|name| of c, theTitle:|name| of c, theSubtitle:subPreFix, theIcon:{theType:missing value, thePath:iconPath}, theAutocomplete:|name| of c, theType:missing value, isValid:"yes", theQuicklook:"/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns", theVars:{{|name|:"subPreFix", value:subPreFix}, {|name|:"argURL", value:theURL}, {|name|:"leaf", value:"false"}, {|name|:"theURL", value:theURL}}, theMods:{cmd:{isValid:"yes", theArg:theURL, theSubtitle:theURL}, alt:missing value, ctrl:missing value, shift:missing value, fn:missing value}, theText:missing value
		end repeat
		
	on error errMsg number errNum
		set schemes to split(theURL, "/")
		set last item of schemes to |name| of n
		set theURL to ""
		
		repeat with x from 1 to count of items of schemes
			if x ≠ (count of items of schemes) then
				set item x of schemes to (item x of schemes) & "/" as text
			end if
			set theURL to theURL & item x of schemes
		end repeat
		
		set theURL to theURL & ".one"
		
		add_result of wf given theUid:(|name| of n), theArg:theURL, theTitle:|name| of n, theSubtitle:theURL, theIcon:{theType:missing value, thePath:"icons/section.png"}, theAutocomplete:|name| of n, theType:missing value, isValid:"yes", theQuicklook:"/Applications/Microsoft OneNote.app/Contents/Resources/OneNote.icns", theVars:{{|name|:"leaf", value:"true"}, {|name|:"theURL", value:theURL}}, theMods:{cmd:{isValid:"yes", theArg:theURL, theSubtitle:theURL}, alt:missing value, ctrl:missing value, shift:missing value, fn:missing value}, theText:missing value
	end try
	
	return wf's to_json("")
end run


on split(theString, theSeparator)
	local saveTID, theResult
	
	set saveTID to AppleScript's text item delimiters
	set AppleScript's text item delimiters to theSeparator
	set theResult to text items of theString
	set AppleScript's text item delimiters to saveTID
	return theResult
end split

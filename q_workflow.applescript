(*
Description:			This AppleScript class provides several useful functions for retrieving, parsing, 
					and formatting data to be used with Alfred 2 Workflow.
Author:				Ursan Razvan
Revised: 			24 March 2013
Version: 			0.5
*)
on new_workflow()
	return my new_workflow_with_bundle(missing value)
end new_workflow
on new_workflow_with_bundle(bundleid)
	script Workflow
		property class : "workflow"
		property _bundle : missing value
		property _cache : missing value
		property _data : missing value
		property _home : missing value
		property _localHash : missing value
		property _name : missing value
		property _path : missing value
		property _preferences : missing value
		property _results : missing value
		property _uuid : missing value
		
		on run {bundleid}
			set _bundle to (system attribute "alfred_workflow_bundleid")
			set _cache to (system attribute "alfred_workflow_cache")
			set _data to (system attribute "alfred_workflow_data")
			set _localHash to (system attribute "alfred_preferences_localhash")
			set _preferences to (system attribute "alfred_preferences")
			set _name to (system attribute "alfred_workflow_name")
			set _uuid to (system attribute "alfred_workflow_uid")
			set _version to (system attribute "alfred_version")
			set _versionBuild to (system attribute "alfred_version_build")
			
			(*
				set my _path to do shell script "pwd"
				if my _path does not end with "/" then set my _path to my _path & "/"
				set my _home to do shell script "printf $HOME"
				set _infoPlist to _path & "info.plist"
				if my q_file_exists(_infoPlist) then
					tell application "System Events"
						tell property list file _infoPlist
							set my _bundle to value of property list item "bundleid" as text
						end tell
					end tell
				end if
				if not my q_is_empty(bundleid) then
					set my _bundle to bundleid
				end if
				set my _cache to (my _home) & "/Library/Caches/com.runningwithcrayons.Alfred-2/Workflow Data/" & (my _bundle) & "/"
				set my _data to (my _home) & "/Library/Application Support/Alfred 2/Workflow Data/" & (my _bundle) & "/"
				if not my q_folder_exists(my _cache) then
					do shell script "mkdir " & quoted form of (my _cache)
				end if
				if not my q_folder_exists(my _data) then
					do shell script "mkdir " & quoted form of (my _data)
				end if
			*)
			set my _results to {}
			return me
		end run
		
		on getPath()
			if my q_is_empty(my _path) then return missing value
			
			return my _path
		end getPath
		
		on getHome()
			if my q_is_empty(my _home) then return missing value
			
			return my _home
		end getHome
		on getResults()
			if my q_is_empty(my _results) then return missing value
			
			return my _results
		end getResults
		on getUID()
			if my q_is_empty(my _uuid) then return missing value
			
			return my _uuid
		end getUID
		on getName()
			if my q_is_empty(my _name) then return missing value
			
			return my _name
		end getName
		on getPreferences()
			if my q_is_empty(my _preferences) then return missing value
			
			return my _preferences
		end getPreferences
		on getBundleID()
			if my q_is_empty(my _bundle) then return missing value
			
			return my _bundle
		end getBundleID
		on getData()
			if my q_is_empty(my _data) then return missing value
			
			return _data
		end getData
		on getCache()
			if my q_is_empty(my _cache) then return missing value
			
			return _cache
		end getCache
		on getVar(varName)
			if my q_is_empty((system attribute varName)) then return missing value
			
			return (system attribute varName) as text
		end getVar
		
		on to_xml(a)
			if (my q_is_empty(a)) and (not my q_is_empty(my _results)) then
				set a to my _results
			else if (my q_is_empty(a)) and (my q_is_empty(my _results)) then
				return missing value
			end if
			set tab2 to tab & tab
			set xml to "<?xml version=\"1.0\"?>" & return & "<items>" & return
			repeat with itemRef in a
				set r to contents of itemRef
				set xml to xml & tab & "<item"
				if not my q_is_empty(theUid of r) then set xml to xml & " uid=\"" & my q_encode(theUid of r) & "\""
				set xml to xml & " arg=\"" & my q_encode(theArg of r) & "\""
				if isValid of r is false then
					set xml to xml & " valid=\"no\""
					if not my q_is_empty(theAutocomplete of r) then
						set xml to xml & " autocomplete=\"" & my q_encode(theAutocomplete of r) & "\""
					end if
				end if
				if not my q_is_empty(theType of r) then
					set xml to xml & " type=\"" & (theType of r) & "\""
				end if
				set xml to xml & ">" & return
				set xml to xml & tab2 & "<title>" & my q_encode(theTitle of r) & "</title>" & return
				set xml to xml & tab2 & "<subtitle>" & my q_encode(theSubtitle of r) & "</subtitle>" & return
				set ic to theIcon of r
				if not my q_is_empty(ic) then
					set xml to xml & tab2 & "<icon"
					if ic starts with "fileicon:" then
						set xml to xml & " type=\"fileicon\""
						set ic to (items 10 thru -1 of ic as text)
					else if ic starts with "filetype:" then
						set xml to xml & " type=\"filetype\""
						set ic to (items 10 thru -1 of ic as text)
					end if
					set xml to xml & ">" & my q_encode(ic) & "</icon>" & return
				else
					if my q_file_exists(my _path & "icon.png") then
						set ic to "icon.png"
						set xml to xml & tab2 & "<icon>icon.png</icon>" & return
					end if
				end if
				set xml to xml & tab & "</item>" & return
			end repeat
			set xml to xml & "</items>"
			return xml
		end to_xml
		
		on to_json(a)
			local r
			local json
			if (my q_is_empty(a)) and (not my q_is_empty(my _results)) then
				set a to my _results
			else if (my q_is_empty(a)) and (my q_is_empty(my _results)) then
				return missing value
			end if
			
			set tab2 to tab & tab
			set tab3 to tab & tab & tab
			set tab4 to tab & tab & tab & tab
			
			set json to "{\"items\": [" & return & return
			repeat with itemRef in a
				set r to contents of itemRef
				set json to json & tab & "{" & return
				if not q_is_empty(theUid of r) then
					set json to json & tab2 & "\"uid\": \"" & encode(theUid of r) & "\"," & return
				end if
				set json to json & tab2 & "\"valid\": \"" & my encode(isValid of r) & "\"," & return
				set json to json & tab2 & "\"title\": \"" & my encode(theTitle of r) & "\"," & return
				set json to json & tab2 & "\"subtitle\": \"" & my encode(theSubtitle of r) & "\"," & return
				set json to json & tab2 & "\"arg\": \"" & my encode(theArg of r) & "\"," & return
				if not q_is_empty(theAutocomplete of r) then
					set json to json & tab2 & "\"autocomplete\": \"" & my encode(theAutocomplete of r) & "\"," & return
				end if
				if not q_is_empty(theQuicklook of r) then
					set json to json & tab2 & "\"quicklookurl\": \"" & my encode(theQuicklook of r) & "\"," & return
				end if
				
				if not q_is_empty(theIcon of r) then
					set ic to theIcon of r
					set json to json & tab2 & "\"icon\": {" & return
					if theType of ic is not "" or theType of ic is not missing value then
						set json to json & tab3 & "\"type\": \"" & my encode(theType of ic) & "\"," & return
					end if
					if thePath of ic is not "" or thePath of ic is not missing value then
						set json to json & tab3 & "\"path\": \"" & my encode(thePath of ic) & "\"" & return
					end if
					set json to json & tab2 & "}," & return
				end if
				
				if not q_is_empty(theVars of r) then
					set vars to ""
					set json to json & tab2 & "\"variables\" : {" & return
					repeat with v in theVars of r
						set varName to |name| of v
						set varVal to value of v
						set json to json & tab3 & "\"" & my encode(varName) & "\": \"" & my encode(varVal) & "\"," & return
					end repeat
					set json to json & tab2 & "}," & return
				end if
				
				
				if not q_is_empty(theMods of r) then
					set json to json & tab2 & "\"mods\": {" & return
					set m to theMods of r
					if not q_is_empty(cmd of m) then
						set json to json & tab3 & "\"cmd\": {" & return
						set json to json & tab4 & "\"valid\": \"" & my encode(isValid of cmd of m) & "\"," & return
						set json to json & tab4 & "\"arg\": \"" & my encode(theArg of cmd of m) & "\"," & return
						set json to json & tab4 & "\"subtitle\": \"" & my encode(theSubtitle of cmd of m) & "\"" & return
						set json to json & tab3 & "}," & return
					end if
					if not q_is_empty(alt of m) then
						set json to json & tab3 & "\"alt\": {" & return
						set json to json & tab4 & "\"valid\": \"" & my encode(isValid of alt of m) & "\"," & return
						set json to json & tab4 & "\"arg\": \"" & my encode(theArg of alt of m) & "\"," & return
						set json to json & tab4 & "\"subtitle\": \"" & my encode(theSubtitle of alt of m) & "\"" & return
						set json to json & tab3 & "}" & return
					end if
					set json to json & tab2 & "}" & return
				end if
				set json to json & tab & "}," & return & return
				
			end repeat
			
			set json to json & return & "]}"
		end to_json
		
		on set_value(a, b, c)
			tell application "System Events"
				if class of a is list then
					set lst to my q_clean_list(a)
					set b to property list file (_get_location of me at b with plist)
					repeat with recordRef in lst
						set r to contents of recordRef
						make new property list item at end of property list items of contents of b ¬
							with properties {kind:(class of (theValue of r)), name:(theKey of r), value:(theValue of r)}
					end repeat
				else
					set c to property list file (_get_location of me at c with plist)
					if class of b is list then
						set x to my q_clean_list(b)
					else
						set x to b
					end if
					make new property list item at end of property list items of contents of c ¬
						with properties {kind:(class of x), name:a, value:x}
				end if
			end tell
		end set_value
		on set_values(a, b)
			return my set_value(a, b, "")
		end set_values
		on get_value(a, b)
			tell application "System Events"
				set b to property list file (_get_location of me at b with plist)
				try
					return value of property list item a of contents of b
				end try
			end tell
			return missing value
		end get_value
		on request(website)
			set agent to "Mozilla/5.0 (compatible; MSIE 7.01; Windows NT 5.0)"
			try
				set theContent to do shell script "curl --silent --show-error --max-redirs 5 --connect-timeout 10 --max-time 10 -L -A '" & agent & "' '" & website & "'"
				return theContent
			end try
			return missing value
		end request
		on request_json(website)
			try
				set jsonHelper to my _path & "bin/q_json.helper"
				set scpt to "tell application \"" & jsonHelper & "\" to fetch JSON from \"" & website & "\""
				set scpt to run script scpt
				if scpt = "" then
					return missing value
				else
					return scpt
				end if
			on error
				return missing value
			end try
		end request_json
		on mdfind(query)
			set output to do shell script "mdfind \"" & query & "\""
			return my q_split(output, return)
		end mdfind
		on write_file(a, b)
			set b to _get_location of me at b without plist
			if class of a is list then
				try
					set a to my q_join(a, return)
				on error
					return false
				end try
			else
				try
					set a to a as text
				on error
					return false
				end try
			end if
			try
				set f to open for access b with write permission
				set eof f to 0
				write a to f as «class utf8»
				close access f
				return true
			on error
				close access b
				return false
			end try
		end write_file
		on read_file(a)
			set a to _get_location of me at a without plist
			try
				set f to open for access a
				set sz to get eof f
				close access f
				if sz = 0 then
					tell application "System Events" to delete file a
					return missing value
				else
					return read a as «class utf8»
				end if
			on error
				close access a
				return missing value
			end try
		end read_file
		-- @description
		-- Helper function that just makes it easier to pass values into a function
		-- and create an array result to be passed back to Alfred
		--
		-- @param $theUid - the uid of the result, should be unique
		-- @param $theArg - the argument that will be passed on
		-- @param $theTitle - The title of the result item
		-- @param $theSubtitle - The subtitle text for the result item
		-- @param $theIcon - the icon to use for the result item
		-- @param $isValid - sets whether the result item can be actioned
		-- @param $theAutocomplete - the autocomplete value for the result item
		-- @return list items to be passed back to Alfred
		--
		on add_result given theUid:_uid, theArg:_arg, theTitle:_title, theSubtitle:_sub, theIcon:_icon, theAutocomplete:_auto, theType:_type, isValid:_valid, theQuicklook:_quicklook, theVars:_vars, theMods:_mods
			if _uid is missing value then set _uid to ""
			if _arg is missing value then set _arg to ""
			if _title is missing value then set _title to ""
			if _sub is missing value then set _sub to ""
			if _icon is missing value then set _icon to ""
			if _auto is missing value then set _auto to ""
			if _type is missing value then set _type to ""
			if _valid is missing value then set _valid to "yes"
			
			set temp to {theUid:_uid, theArg:_arg, theTitle:_title, theSubtitle:_sub, theIcon:_icon, theAutocomplete:_auto, theType:_type, isValid:_valid, theQuicklook:_quicklook, theVars:_vars, theMods:_mods}
			if my q_is_empty(_type) then
				set temp's theType to missing value
			end if
			set end of (my _results) to temp
			return temp
		end add_result
		on _make_plist(plistPath)
			tell application "System Events"
				set parentElement to make new property list item with properties {kind:record}
				set plistFile to ¬
					make new property list file with properties {contents:parentElement, name:plistPath}
			end tell
			return plistFile
		end _make_plist
		on _get_location at pathOrName given plist:isPlist
			if pathOrName is missing value or my q_is_empty(pathOrName) then set pathOrName to "settings.plist"
			if my q_file_exists(pathOrName) then
			else if my q_file_exists(my _path & pathOrName) then
				set location to my _path & pathOrName
			else if my q_file_exists(my _data & pathOrName) then
				set location to my _data & pathOrName
			else if my q_file_exists(my _cache & pathOrName) then
				set location to my _cache & pathOrName
			else
				set location to my _data & pathOrName
				if isPlist then
					my _make_plist(location)
				else
					try
						set f to open for access location with write permission
						set eof of f to 0
						close access f
					on error
						do shell script "touch " & quoted form of location
					end try
				end if
			end if
			return location
		end _get_location
	end script
	return run script Workflow with parameters {bundleid}
end new_workflow_with_bundle
on q_join(l, delim)
	if class of l is not list or l is missing value then return ""
	repeat with i from 1 to length of l
		if item i of l is missing value then
			set item i of l to ""
		end if
	end repeat
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delim
	set output to l as text
	set AppleScript's text item delimiters to oldDelims
	return output
end q_join
on q_split(s, delim)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to delim
	set output to text items of s
	set AppleScript's text item delimiters to oldDelims
	return output
end q_split
on q_file_exists(theFile)
	if my q_path_exists(theFile) then
		tell application "System Events"
			return (class of (disk item theFile) is file)
		end tell
	end if
	return false
end q_file_exists
on q_folder_exists(theFolder)
	if my q_path_exists(theFolder) then
		tell application "System Events"
			return (class of (disk item theFolder) is folder)
		end tell
	end if
	return false
end q_folder_exists
on q_path_exists(thePath)
	if thePath is missing value or my q_is_empty(thePath) then return false
	try
		if class of thePath is alias then return true
		if thePath contains ":" then
			alias thePath
			return true
		else if thePath contains "/" then
			POSIX file thePath as alias
			return true
		else
			return false
		end if
	on error msg
		return false
	end try
end q_path_exists
on q_is_empty(str)
	if str is missing value then return true
	return length of (my q_trim(str)) is 0
end q_is_empty
on q_trim(str)
	if class of str is not text or class of str is not string or str is missing value then return str
	if str is "" then return str
	repeat while str begins with " "
		try
			set str to items 2 thru -1 of str as text
		on error msg
			return ""
		end try
	end repeat
	repeat while str ends with " "
		try
			set str to items 1 thru -2 of str as text
		on error
			return ""
		end try
	end repeat
	return str
end q_trim
on q_clean_list(lst)
	if lst is missing value or class of lst is not list then return lst
	set l to {}
	repeat with lRef in lst
		set i to contents of lRef
		if i is not missing value then
			if class of i is not list then
				set end of l to i
			else if class of i is list then
				set end of l to my q_clean_list(i)
			end if
		end if
	end repeat
	return l
end q_clean_list
on q_encode(str)
	if class of str is not text or my q_is_empty(str) then return str
	set s to ""
	repeat with sRef in str
		set c to contents of sRef
		if c is in {"&", "'", "\"", "<", ">", tab} then
			if c is "&" then
				set s to s & "&amp;"
			else if c is "'" then
				set s to s & "&apos;"
			else if c is "\"" then
				set s to s & "&quot;"
			else if c is "<" then
				set s to s & "&lt;"
			else if c is ">" then
				set s to s & "&gt;"
			else if c is tab then
				set s to s & "&#009;"
			end if
		else
			set s to s & c
		end if
	end repeat
	return s
end q_encode
on q_date_to_unixdate(theDate)
	set {day:d, year:y, time:t} to theDate
	copy theDate to b
	set b's month to January
	set m to (b - 2500000 - theDate) div -2500000
	tell (y * 10000 + m * 100 + d) as text
		set UnixDate to text 5 thru 6 & "/" & text 7 thru 8 & "/" & text 1 thru 4
	end tell
	set h24 to t div hours
	set h12 to (h24 + 11) mod 12 + 1
	if (h12 = h24) then
		set ampm to " AM"
	else
		set ampm to " PM"
	end if
	set min to t mod hours div minutes
	set s to t mod minutes
	tell (1000000 + h12 * 10000 + min * 100 + s) as text
		set UnixTime to text 2 thru 3 & ":" & text 4 thru 5 & ":" & text 6 thru 7 & ampm
	end tell
	return UnixDate & " " & UnixTime
end q_date_to_unixdate
on q_unixdate_to_date(theUnixDate)
	return date theUnixDate
end q_unixdate_to_date
on q_timestamp_to_date(timestamp)
	if length of timestamp = 13 then
		set timestamp to characters 1 thru -4 of timestamp as text
	end if
	set h to do shell script "date -r " & timestamp & " \"+%Y %m %d %H %M %S\""
	set mydate to current date
	set year of mydate to (word 1 of h as integer)
	set month of mydate to (word 2 of h as integer)
	set day of mydate to (word 3 of h as integer)
	set hours of mydate to (word 4 of h as integer)
	set minutes of mydate to (word 5 of h as integer)
	set seconds of mydate to (word 6 of h as integer)
	return mydate
end q_timestamp_to_date
on q_date_to_timestamp(theDate)
	return ((current date) - (date ("1/1/1970")) - (time to GMT)) as miles as text
end q_date_to_timestamp
on q_send_notification(theMessage, theDetails, theExtra)
	set _path to do shell script "pwd"
	if _path does not end with "/" then set _path to _path & "/"
	if theMessage is missing value then set theMessage to ""
	if theDetails is missing value then set theDetails to ""
	if theExtra is missing value then set theExtra to ""
	if my q_trim(theMessage) is "" and my q_trim(theExtra) is "" then set theMessage to "notification"
	try
		do shell script (quoted form of _path & "bin/q_notifier.helper com.runningwithcrayons.Alfred-2 " & quoted form of theMessage & " " & quoted form of theDetails & " " & quoted form of theExtra)
	end try
end q_send_notification
on q_notify()
	my q_send_notification("", "", "")
end q_notify
on q_encode_url(str)
	local str
	try
		return (do shell script "/bin/echo " & quoted form of str & ¬
			" | perl -MURI::Escape -lne 'print uri_escape($_)'")
	on error
		return missing value
	end try
end q_encode_url
on q_decode_url(str)
	local str
	try
		return (do shell script "/bin/echo " & quoted form of str & ¬
			" | perl -MURI::Escape -lne 'print uri_unescape($_)'")
	on error
		return missing value
	end try
end q_decode_url

on encode(theString)
	if class of theString is not text then return theString
	if theString contains "\"" then set theString to my SearchandReplace(theString, "\"", "\\\"")
	--if theString contains "\\" then set theString to my SearchandReplace(theString, "\\", "\\\\")
	return theString
end encode

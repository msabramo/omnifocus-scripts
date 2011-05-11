tell application "OmniFocus"
	tell front document
		set destinationFolder to my getFolderFromUser()
		
		repeat with anItem in my getItems()
			move anItem to after last project of folder id destinationFolder
		end repeat
	end tell
end tell

property pstrDBPath : "~/Library/Caches/com.omnigroup.OmniFocus/OmniFocusDatabase2"
property pstrFieldDelimiter : "~|~"
property plstObjects : {}
property pSearch : ""

on getFolderFromUser()
	set strSearch to text returned of (display dialog "Enter a phrase in the name of the folder" default answer pSearch)
	set pSearch to strSearch
	
	set blnCreated to false
	set blnFound to false
	
	-- LOOP UNTIL SOMETHING FOUND OR CREATED, OR UNTIL THE USER TAPS "ESC"
	repeat while (blnCreated is false and blnFound is false)
		set strQuery to "select t.name from folder t where t.name like \"%" & EscapeQuote(strSearch) & "%\" order by t.name;"
		--tell application "Finder" to set the clipboard to strQuery
		set plstObjects to {}
		-- try
		set plstObjects to paragraphs of runquery(strQuery)
		(*
		on error
		display dialog "The SQL schema for the OmniFocus cache may have changed in a recent update of OF." & return & return & Â
			"Look on the OmniFocus user forums for an updated version of this script." buttons {"OK"} default button {"OK"}
		return
		end try
		*)
		
		if plstObjects is {} then -- NO MATCHFOUND 
			-- OFFER: ESC, CREATE FOLDER, OR RUN AMENDED SEARCH
			set varResponse to display dialog "No matches found ..." & Â
				return & return & "Try a modified search, or create a folder of this name ?" default answer strSearch buttons {"Esc", pCreateProj, pReSearch} default button {pReSearch} with title pTitle & "  Ver " & pVersion
			set strButton to button returned of varResponse
			if strButton is pCreateProj then
				tell application id "com.omnigroup.omnifocus"
					tell default document
						set oProj to make new folder with properties {name:strSearch}
						set oWin to make new document window
						tell oWin
							set focus to {oProj}
							tell sidebar
								select {oProj}
							end tell
							tell content
								select {oProj}
							end tell
						end tell
						set oTask to make new task at end of tasks of oProj
						tell oWin to tell content to select {oTask}
					end tell
				end tell
				tell application id "com.apple.finder" to activate front window of application id "com.omnigroup.omnifocus"
				return
			else if strButton is pReSearch then
				set strSearch to text returned of varResponse
			else
				return
			end if
		else -- At least one match found
			set blnFound to true
			set lngMatches to length of plstObjects
			if lngMatches > 1 then
				set varChoice to choose from list plstObjects with prompt (lngMatches as string) & " folders where name contains " & quote & strSearch & quote & return & return & Â
					"Select one:" default items {first item of plstObjects}
				if varChoice is false then return
				varChoice
				
				set strQuery to "select persistentidentifier from folder where "
				repeat with oChoice in varChoice
					set strQuery to strQuery & "name like \"" & EscapeQuote(oChoice) & "\"" & " or "
				end repeat
				set strQuery to text 1 thru -5 of strQuery
				set strQuery to strQuery & ";"
				
				set plstObjects to first item of paragraphs of runquery(strQuery)
			else
				set plstObjects to runquery("select persistentidentifier from folder where name like \"" & EscapeQuote(first item of plstObjects) & "\";")
			end if
		end if
	end repeat
	
	return plstObjects
	
	tell application "OmniFocus"
		tell front document
		end tell
	end tell
end getFolderFromUser

on runquery(strQuery)
	set strCmd to "sqlite3 -separator '" & pstrFieldDelimiter & "' " & pstrDBPath & space & quoted form of strQuery
	-- tell application "Finder" to set the clipboard to strQuery
	do shell script strCmd
end runquery

on EscapeQuote(strSearch)
	FindReplace(strSearch, "\"", "\"\"")
end EscapeQuote

on FindReplace(strText, strFind, strReplace)
	if the strText contains strFind then
		set AppleScript's text item delimiters to strFind
		set lstParts to text items of strText
		set AppleScript's text item delimiters to strReplace
		set strText to lstParts as string
		set AppleScript's text item delimiters to space
	end if
	return strText
end FindReplace

on getItems()
	set myItems to {}
	
	tell application "OmniFocus"
		tell front document
			tell sidebar of document window 1
				set theSelectedItems to value of every selected tree
				repeat with anItem in theSelectedItems
					copy anItem to the end of myItems
				end repeat
			end tell
		end tell
	end tell
	
	return myItems
end getItems
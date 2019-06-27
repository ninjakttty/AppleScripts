(*
===============================================================================
                        Import WFO Schedule to Calendar
===============================================================================

Version: 1.1                                     Updated: 06/11/19, 01:54:40 AM
By: Kevin Funderburg

PURPOSE:

This script is designed to gather an employees schedule data from WFO so the
schedule can be copied into a personal calendar so employees can see their

work schedule on their personal devices easily.
NOTES:

• IN ORDER FOR THIS TO WORK, you must have already logged into WFO.
The script basically reads the data from the Schedule Editor page of WFO and
since I don't know your login credentials, I can't login to WFO through the
script.

• The Segment button needs to be clicked at the top of the Schedule Editor.
Part of the script does that for you, but for those who like the Resolution
view more, they will see that switch, no biggy.

• For right now, this only works 1 week at a time, so make sure at the top of
Schedule Editor that the "Week" button is checked, not "Day" or "Month"
    • I will expand this in the future


REQUIRED:
	1.	Mac OS X Yosemite 10.10.5+
	2.	Mac Applications
			• Safari

	3.	EXTERNAL OSAX Additions/LIBRARIES/FUNCTIONS
			• None

VERSION HISTORY:
1.0 - Initial version.
1.1 - Changed to applet for easier distribution
===============================================================================
*)
use AppleScript version "2.4" -- Yosemite (10.10) or later
use framework "Foundation"
use framework "EventKit"
use scripting additions

-- classes, constants, and enums used
property NSDictionary : a reference to current application's NSDictionary
property NSString : a reference to current application's NSString
property cal : missing value
property scheduleEditorURL : "[INSERT URL HERE]" # I removed the actual URL here for confidentiality reasons

on run
	try
		setCalendar()
		set progress total steps to -1 -- indeterminant progress
		set progress completed steps to 0
		set progress description to "Getting WFO schedule data"
		set progress additional description to "Waiting for WFO to load..."

		set theText to getWFOtext()
		set scheduleData to parseWFOtext(theText)
		--set cal to "Work"
		set progress completed steps to 1
		set progress description to "Deleting the events of week " & _date of item 1 of scheduleData

		-- Delete all calendar items for this week to avoid duplicates
		set weekStartDate to date (_date of item 1 of scheduleData)
		deleteCalEvents(weekStartDate)

		set progress description to "Creating events of week " & _date of item 1 of scheduleData
		set progress total steps to 7
		set counter to 0
		repeat with d in scheduleData
			set counter to counter + 1
			set progress completed steps to counter
			set progress additional description to "Day " & counter & " of 7"

			if item 1 of d's segments ≠ "No data" then

				repeat with s in segments of d
					log s
					if s does not start with "ACPHON" then
						set segmentName to item 1 of my SearchWithRegEx("^.* - (.*) (\\d+:\\d+ [A|P]M) - (\\d+:\\d+ [A|P]M)", s, 1)
						if segmentName = "Shift (container)" then set segmentName to "Shift"
						set startTime to date (_date of d & " at " & item 1 of my SearchWithRegEx("^.* - (.*) (\\d+:\\d+ [A|P]M) - (\\d+:\\d+ [A|P]M)", s, 2))
						set endTime to date (_date of d & " at " & my SearchWithRegEx("^.* - (.*) (\\d+:\\d+ [A|P]M) - (\\d+:\\d+ [A|P]M)", s, 3))
						makeCalEvents(segmentName, startTime, endTime)
					end if
				end repeat
			end if

		end repeat

		set progress total steps to -1 -- indeterminant progress
		set progress completed steps to 0
		set progress description to "Import complete!"
		set progress additional description to ""
		display notification "Calendar imported successfully!" with title "WFO to Calendar"
		delay 2
	on error errMsg number errNum from errFrom to errTo partial result errPartialResult
		set progress total steps to -1
		set progress completed steps to 0
		set progress additional description to "Stopping..."
		error errMsg number errNum from errFrom to errTo partial result errPartialResult -- resignal the error

	end try
end run


-- @description
-- checks if a calendar has been chosen to put the events to and if not will prompt
-- with a list of the calendars to choose from
--
-- @return none
--
on setCalendar()
	tell application "Calendar" to set cals to name of every calendar
	set p to POSIX path of (path to me) & "Contents/Resources/data.plist"
	tell application "System Events"
		set info to value of property list file p
		set cal to theCalendar of info
	end tell

	if cals does not contain cal or cal is missing value then
		choose from list cals ¬
			with title ¬
			"WFO to Calendar" with prompt ¬
			"Choose the calendar to import to" default items 1 ¬
			OK button name ¬
			"OK" cancel button name ¬
			"Cancel" multiple selections allowed false ¬
			without empty selection allowed

		set cal to item 1 of result
		tell application "System Events"
			tell property list file p
				set value of property list item "theCalendar" to cal
			end tell
		end tell
	end if

end setCalendar


-- @description
-- This checks if the schdule editor is already open in Safari. This script
-- depends on the schedule editor to be open or it won't be able to get the
-- schedule data. If it is not open, it will ask the user to log into WFO
-- and try again. Once the schedule editor is open, it then gathers the
-- schedule text to be parsed later.
--
-- @return text of Safari's tab
--
on getWFOtext()
	tell application "Safari"
		tell front window

			set URLs to URL of every tab

			-- Check if WFO is already open so another tab isn't opened
			if URLs contains scheduleEditorURL then
				repeat with n from 1 to (count of URLs)
					if item n of URLs is scheduleEditorURL then
						set URLfound to true
						exit repeat
					end if
				end repeat
				set thetab to tab n
			else
				open location scheduleEditorURL
				set thetab to current tab
			end if

			my waitForSafariToLoad("Saturday", thetab) -- Pause until schedule has loaded
			delay 0.2
			set theText to text of thetab
			if theText does not contain "Clear Day" then -- Means the Resolution button is clicked
				-- click the Segment button in Schedule Editor
				my doJava:"click" onType:"class" withIdentifier:"toolbar-item-caption" withElementNum:"0" withSetValue:(missing value) inTab:thetab
				my waitForSafariToLoad("Clear Day", thetab) -- Pause again in case the Segment view needs to load
				delay 1.5
				set theText to text of thetab
			end if
		end tell
	end tell

	return theText
end getWFOtext


-- @description
-- Creates the events to be created in the calendar designated by the
-- setCalendar() function.
--
-- @param $theTitle - description of the shift segment
-- @param $startDate - start time of the shift segment
-- @param $endDate - end time of the shift segment
-- @return none
--
on makeCalEvents(theTitle, startDate, endDate)
	set eventStartDate to current application's NSDate's dateWithTimeInterval:0.0 sinceDate:startDate
	set eventEndDate to current application's NSDate's dateWithTimeInterval:0.0 sinceDate:endDate
	set listOfCalNames to {cal} -- list of one or more calendar names

	set listOfCalTypes to {1} -- list of one or more calendar types: : Local = 0, CalDAV/iCloud = 1, Exchange = 2, Subscription = 3, Birthday = 4
	-- create event store and get the OK to access Calendars
	set theEKEventStore to current application's EKEventStore's alloc()'s init()
	theEKEventStore's requestAccessToEntityType:0 completion:(missing value)

	-- check if app has access; this will still occur the first time you OK authorization
	set authorizationStatus to current application's EKEventStore's authorizationStatusForEntityType:0 -- work around enum bug
	if authorizationStatus is not 3 then
		display dialog "Access must be given in System Preferences" & linefeed & "-> Security & Privacy first." buttons {"OK"} default button 1
		tell application "System Preferences"
			activate
			tell pane id "com.apple.preference.security" to reveal anchor "Privacy"
		end tell
		error number -128
	end if

	-- get calendars that can store events
	set theCalendars to theEKEventStore's calendarsForEntityType:0
	-- filter out the one you want
	set theNSPredicate to current application's NSPredicate's predicateWithFormat_("title IN %@ AND type IN %@", listOfCalNames, listOfCalTypes)
	set calsToSearch to theCalendars's filteredArrayUsingPredicate:theNSPredicate
	if (count of calsToSearch) < 1 then error "No such calendar(s)."

	set ev to current application's EKEvent's eventWithEventStore:theEKEventStore
	ev's setCalendar:(item 1 of calsToSearch)
	ev's setTitle:theTitle
	ev's setStartDate:eventStartDate
	ev's setEndDate:eventEndDate

	set {theResult, theError} to theEKEventStore's saveEvent:ev span:(current application's EKSpanThisEvent) commit:true |error|:(reference)
	if not theResult as boolean then error (theError's |localizedDescription|() as text)
end makeCalEvents


-- @description
-- Deletes all the events of the designated calendar to avoid duplicate
-- events being created.
--
-- @param $weekStartDate - First day of the week the events will be created in
-- @return none
--
on deleteCalEvents(weekStartDate)
	set weekEndDate to weekStartDate + (7 * days)
	set listOfCalNames to {cal} -- list of one or more calendar names
	set listOfCaTypes to {1} -- list of one or more calendar types: : Local = 0, CalDAV/iCloud = 1, Exchange = 2, Subscription = 3, Birthday = 4
	-- create start date and end date for occurances
	set nowDate to current application's NSDate's |date|()
	set weekStartDate to current application's NSDate's dateWithTimeInterval:0.0 sinceDate:weekStartDate
	--set weekStartDate to current application's NSCalendar's currentCalendar()'s dateBySettingHour:0 minute:0 |second|:0 ofDate:weekStartDate options:0
	set todaysDate to current application's NSCalendar's currentCalendar()'s dateBySettingHour:0 minute:0 |second|:0 ofDate:nowDate options:0
	set weekEndDate to weekStartDate's dateByAddingTimeInterval:7 * days

	-- create event store and get the OK to access Calendars
	set theEKEventStore to current application's EKEventStore's alloc()'s init()
	theEKEventStore's requestAccessToEntityType:0 completion:(missing value)

	-- check if app has access; this will still occur the first time you OK authorization
	set authorizationStatus to current application's EKEventStore's authorizationStatusForEntityType:0 -- work around enum bug
	if authorizationStatus is not 3 then
		display dialog "Access must be given in System Preferences" & linefeed & "-> Security & Privacy first." buttons {"OK"} default button 1
		tell application "System Preferences"
			activate
			tell pane id "com.apple.preference.security" to reveal anchor "Privacy"
		end tell
		error number -128
	end if

	-- get calendars that can store events
	set theCalendars to theEKEventStore's calendarsForEntityType:0
	-- filter out the one you want
	set theNSPredicate to current application's NSPredicate's predicateWithFormat_("title IN %@ AND type IN %@", listOfCalNames, listOfCaTypes)
	set calsToSearch to theCalendars's filteredArrayUsingPredicate:theNSPredicate
	if (count of calsToSearch) < 1 then error "No such calendar(s)."

	-- find matching events
	set thePred to theEKEventStore's predicateForEventsWithStartDate:weekStartDate endDate:weekEndDate calendars:calsToSearch
	set theEvents to (theEKEventStore's eventsMatchingPredicate:thePred)
	-- sort by date
	set theEvents to theEvents's sortedArrayUsingSelector:"compareStartDateWithEvent:"

	repeat with e in theEvents
		(theEKEventStore's removeEvent:e span:(current application's EKSpanThisEvent) commit:true |error|:(missing value))
	end repeat

end deleteCalEvents


-- @description
-- Parse the data created by getWFOtext() into a record of shift segments for
-- each work day
--
-- @param $wfoText - Text of schedule editor webpage
-- @return record of shift segments
--
on parseWFOtext(wfoText)
	set daysofweek to {"Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"}
	set scheduleData to {}
	set dataScraped to false
	set n to 1

	repeat until dataScraped
		set p to paragraph n of wfoText
		log p
		if (count of words of p) > 0 then
			if first word of p is in daysofweek then
				set theday to first word of p
				set thedate to my extractBetween(p, ", ", return)
				set datastart to n
				set segdata to {}
				set segmentend to false

				repeat
					set n to n + 1
					log paragraph n of wfoText
					if word 1 of paragraph n of wfoText is not in daysofweek and paragraph n of wfoText is not "YOUR SCHEDULE EDITS" then
						if paragraph n of wfoText does not contain "Clear day" then
							set end of segdata to paragraph n of wfoText
						end if
					else
						exit repeat
					end if
				end repeat
				set n to n - 1
				set end of scheduleData to {_dow:theday, _date:thedate, segments:segdata}
			end if
		end if

		if paragraph n of wfoText is "Loading complete" then
			set dataScraped to true
		else
			set n to n + 1
		end if

	end repeat

	return scheduleData
end parseWFOtext


-- @description
-- Delay script from proceeding until Safari has loaded the page completely
--
-- @param $SearchText - text that only appears in the webpage when it is loaded
-- @param $thetab - the tab that contains the text to searched for
-- @return none
--
on waitForSafariToLoad(SearchText, thetab)
	set tabText to ""
	set failsafe to 0

	tell application "Safari"
		tell front window

			repeat until tabText contains SearchText
				set tabText to text of thetab
				delay 0.1
				set failsafe to failsafe + 1
				if failsafe = 300 then
					error "Script timed out, Schedule Editor not found" & return & "Login to WFO and click the Schedule Editor link and try again" number -100
				end if
			end repeat

		end tell
	end tell
end waitForSafariToLoad


-- @Description
-- Performs basic JavaScript actions in Safari
--
-- @param - theAction: get, set, click, submit
-- @param - theType: id, class, etc
-- @param - id: the string associated with the object
-- @param - num: the index of the object
-- @param - theValue: value an object is to be set to (set to missing value by default)
-- @param - theTab: the tab the action is to be performed in
-- @return none
--
on doJava:theAction onType:theType withIdentifier:theID withElementNum:num withSetValue:theValue inTab:thetab
	if theType = "id" then
		set getBy to "getElementById"
		set theJavaEnd to ""
	else
		if theType = "class" then
			set theType to "ClassName"
		else if theType = "name" then
			set theType to "Name"
		else if theType = "tag" then
			set theType to "TagName"
		end if
		set getBy to "getElementsBy" & theType
		set theJavaEnd to "[" & num & "]"
	end if

	if theAction = "click" then
		set theJavaEnd to theJavaEnd & ".click();"
	else if theAction = "get" then
		set theJavaEnd to theJavaEnd & ".innerHTML;"
	else if theAction = "set" then
		set theJavaEnd to theJavaEnd & ".value ='" & theValue & "';"
	else if theAction = "submit" then
	else if theAction = "force" then
	end if

	set theJava to "document." & getBy & "('" & theID & "')" & theJavaEnd

	tell application "Safari"
		if thetab is missing value then set thetab to front document
		tell thetab
			if theAction = "get" then
				set input to do JavaScript theJava
				return input
			else
				do JavaScript theJava
			end if
		end tell
	end tell

end doJava:onType:withIdentifier:withElementNum:withSetValue:inTab:

-- @description
-- Handler for regular expression searching
--
-- @param $thePattern - regex pattern
-- @param $theString - the string to search
-- @param $n - capturing group
-- @return matching string
--
on SearchWithRegEx(thePattern, theString, n)
	set theNSString to NSString's stringWithString:theString
	set theOptions to ((current application's NSRegularExpressionDotMatchesLineSeparators) as integer) + ((current application's NSRegularExpressionAnchorsMatchLines) as integer)
	set theRegEx to current application's NSRegularExpression's regularExpressionWithPattern:thePattern options:theOptions |error|:(missing value)
	set theFinds to theRegEx's matchesInString:theNSString options:0 range:{location:0, |length|:theNSString's |length|()}
	set theResult to {}
	repeat with i from 1 to count of items of theFinds
		set oneFind to (item i of theFinds)
		if (oneFind's numberOfRanges()) as integer < (n + 1) then
			set end of theResult to missing value
		else
			set theRange to (oneFind's rangeAtIndex:n)
			set end of theResult to (theNSString's substringWithRange:theRange) as string
		end if
	end repeat
	return theResult
end SearchWithRegEx


-- @description
-- Extract text between 2 delimiters
--
-- @param $SearchText - text to search
-- @param $startText - text before
-- @param $endText - text after
-- @return text between
--
on extractBetween(SearchText, startText, endText)
	set tid to AppleScript's text item delimiters
	set AppleScript's text item delimiters to startText
	set endItems to text of text item -1 of SearchText
	set AppleScript's text item delimiters to endText
	set beginningToEnd to text of text item 1 of endItems
	set AppleScript's text item delimiters to tid
	return beginningToEnd
end extractBetween

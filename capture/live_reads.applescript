-- Open a workbook (argument 1) and read it while it is open: each further argument is an
-- expression of Excel's AppleScript dictionary (such as: font size of font object of range "A1"
-- of active sheet), evaluated with "get" in turn. Returns each result as text, tab-separated;
-- an expression Excel cannot answer gives "error: " and its message. Closes the workbook
-- without saving. Non-interactive: alerts are off while it runs and restored on every exit
-- path. Touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set tb to tab -- the separator, taken before the tell block
	set answers to {}
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			repeat 20 times
				if exists workbook xlName then exit repeat
				delay 0.25
			end repeat
			activate object workbook xlName
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	repeat with k from 2 to count of argv
		set expr to item k of argv
		try
			set answer to run script ("tell application \"Microsoft Excel\" to get (" & expr & ")")
			set end of answers to my shown(answer)
		on error errMsg
			set end of answers to "error: " & errMsg
		end try
	end repeat
	tell application "Microsoft Excel"
		try
			close workbook xlName saving no
		end try
		set display alerts to true
	end tell
	set AppleScript's text item delimiters to tb
	set out to answers as text
	set AppleScript's text item delimiters to ""
	return out
end run

on shown(answer)
	if answer is missing value then return "missing value"
	if class of answer is list then
		set parts to {}
		repeat with x in answer
			set end of parts to my shown(contents of x)
		end repeat
		set AppleScript's text item delimiters to "|"
		set joined to parts as text
		set AppleScript's text item delimiters to ""
		return joined
	end if
	try
		return (class of answer as text) & ":" & (answer as text)
	on error
		return (class of answer as text) & ":?"
	end try
end shown

-- Open a workbook (argument 1), return what A1:E1 of its first sheet show (their string
-- values, as Excel draws them, and the class of each value), then, when argument 2 is given,
-- save it as an .xlsx copy there; close it without saving. Non-interactive: alerts are off while
-- it runs and restored on every exit path. Touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopy to ""
	if (count of argv) > 1 then set xlCopy to item 2 of argv
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			repeat 20 times
				if exists workbook xlName then exit repeat
				delay 0.25
			end repeat
			set xlWb to workbook xlName
			set xlSheet to worksheet 1 of xlWb
			set shown to {}
			repeat with cellRef in {"A1", "B1", "C1", "D1", "E1"}
				set xlCell to range (cellRef as text) of xlSheet
				set t to string value of xlCell
				try
					set k to (class of (value of xlCell)) as text
				on error
					set k to "?"
				end try
				set end of shown to t & tab & k
			end repeat
			if xlCopy is not "" then
				set xlCopyName to do shell script "basename " & quoted form of xlCopy
				save workbook as xlWb filename xlCopy file format workbook normal file format
				close workbook xlCopyName saving no
			else
				close xlWb saving no
			end if
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	set AppleScript's text item delimiters to tab
	set out to shown as text
	set AppleScript's text item delimiters to ""
	return out
end run

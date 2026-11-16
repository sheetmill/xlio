-- Save a workbook that is already open in Excel (argument 1, its name) as an .xlsx copy
-- (argument 2; an .xlsm copy when argument 3 is "xlsm"), then close both without saving. For a file opened the way a person opens it
-- (Launch Services); the open itself is not done here. Alerts are off while it runs and
-- restored on every exit path.
on run argv
	set xlName to item 1 of argv
	set xlCopy to item 2 of argv
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set macroCopy to (count of argv) > 2 and item 3 of argv is "xlsm"
	tell application "Microsoft Excel"
		if not (exists workbook xlName) then error "not open in Excel: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			if macroCopy then
				save workbook as (workbook xlName) filename xlCopy file format macro enabled XML file format
			else
				save workbook as (workbook xlName) filename xlCopy file format workbook normal file format
			end if
			close workbook xlCopyName saving no
			set display alerts to true
		on error errMsg number errNum
			try
				close workbook xlName saving no
			end try
			try
				close workbook xlCopyName saving no
			end try
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return "ok"
end run

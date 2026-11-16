-- Save a workbook that is already open in Excel (argument 1, its name) as an .xlsx copy
-- (argument 2), then close both without saving. For a file opened the way a person opens it
-- (Launch Services); the open itself is not done here. Alerts are off while it runs and
-- restored on every exit path.
on run argv
	set xlName to item 1 of argv
	set xlCopy to item 2 of argv
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	tell application "Microsoft Excel"
		if not (exists workbook xlName) then error "not open in Excel: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			save workbook as (workbook xlName) filename xlCopy file format workbook normal file format
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

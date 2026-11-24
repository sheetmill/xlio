-- Open a text file (argument 1) in Excel for Mac, save the workbook Excel makes of it as an .xlsx
-- copy (argument 2), close it. Excel may name the workbook otherwise than the file (a[b]c.csv
-- opens as a(b)c.csv), so the workbook is the one the open adds. Non-interactive: alerts are off
-- while it runs and restored on every exit path. Touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlCopy to item 2 of argv
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	tell application "Microsoft Excel"
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set priorNames to name of every workbook
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set opened to my newNames(priorNames)
			if (count of opened) is 0 then error "no workbook" number -50
			set xlWb to workbook (item 1 of opened)
			save workbook as xlWb filename xlCopy file format workbook normal file format
			close workbook xlCopyName saving no
			my closeNew(priorNames)
			set display alerts to true
		on error errMsg number errNum
			my closeNew(priorNames)
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
	return "ok"
end run

on newNames(priorNames)
	tell application "Microsoft Excel"
		set found to {}
		repeat with xlName in (name of every workbook)
			if (xlName as text) is not in priorNames then set end of found to (xlName as text)
		end repeat
		return found
	end tell
end newNames

on closeNew(priorNames)
	tell application "Microsoft Excel"
		repeat with xlName in my newNames(priorNames)
			try
				close workbook (xlName as text) saving no
			end try
		end repeat
	end tell
end closeNew

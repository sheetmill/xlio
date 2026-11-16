-- Open one workbook (argument 1), read A1 of its first sheet, the number of sheets, the widths of
-- columns A and B, the heights of rows 1 and 2 and A1's font size, close it without saving.
-- Returns "opened" and those, tab-separated, or "refused<TAB>error number". A workbook
-- that Excel opens under another name (a repaired copy) is reported and closed too: it is this
-- script's own. Non-interactive: alerts are off while it runs and restored on every exit path.
-- Touches no other open workbook; uses no clipboard.
on run argv
	set xlPath to item 1 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set tabChar to ASCII character 9
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		set priorNames to name of every workbook
		set display alerts to false
		try
			try
				open workbook workbook file name xlPath
			on error errMsg number errNum
				set answer to "refused" & tabChar & (errNum as text)
				my closeNew(priorNames)
				set display alerts to true
				return answer
			end try
			if exists workbook xlName then
				set opened to {xlName}
			else
				set opened to my newNames(priorNames)
			end if
			if (count of opened) is 0 then
				set display alerts to true
				return "refused" & tabChar & "no workbook"
			end if
			set xlWb to workbook (item 1 of opened)
			set xlValue to value of range "A1" of worksheet 1 of xlWb
			set xlSheets to count of worksheets of xlWb
			set xlRange to range "A1" of worksheet 1 of xlWb
			set xlWidth to column width of xlRange
			set xlHeight to row height of xlRange
			set xlWidthB to column width of range "B1" of worksheet 1 of xlWb
			set xlHeight2 to row height of range "A2" of worksheet 1 of xlWb
			set xlFont to font size of font object of xlRange
			set answer to "opened" & tabChar & (xlValue as text) & tabChar & (xlSheets as text) & tabChar & (xlWidth as text) & tabChar & (xlHeight as text) & tabChar & (xlWidthB as text) & tabChar & (xlHeight2 as text) & tabChar & (xlFont as text)
			close xlWb saving no
			my closeNew(priorNames)
			set display alerts to true
			return answer
		on error errMsg number errNum
			my closeNew(priorNames)
			set display alerts to true
			error errMsg number errNum
		end try
	end tell
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

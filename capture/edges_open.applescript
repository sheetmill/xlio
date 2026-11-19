-- Open a workbook (argument 1), read the widths of columns A and B and A1's font size, save it
-- as an .xlsx workbook under another name (argument 2) and close that copy. Returns those three
-- values, tab-separated. Non-interactive: alerts are off while it runs and restored on every
-- exit path. Touches no other open workbook; uses no clipboard.
on run argv
	set xlPath to item 1 of argv
	set xlCopy to item 2 of argv
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set tabChar to ASCII character 9
	tell application "Microsoft Excel"
		-- A workbook of this name already open in Excel is someone's: leave it and stop.
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			set xlWb to workbook xlName
			set xlSheet to worksheet 1 of xlWb
			set xlWidthA to column width of range "A1" of xlSheet
			set xlWidthB to column width of range "B1" of xlSheet
			set xlFont to font size of font object of range "A1" of xlSheet
			set answer to (xlWidthA as text) & tabChar & (xlWidthB as text) & tabChar & (xlFont as text)
			save workbook as xlWb filename xlCopy file format workbook normal file format
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
	return answer
end run

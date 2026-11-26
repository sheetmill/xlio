-- Open a text file (argument 1) in Excel for Mac and read cells A1:A<n> (argument 3, n) while it
-- is open: each cell's string value (the text Excel draws) and the class of its value. Then save
-- the workbook as an .xlsx copy (argument 2), write =ERROR.TYPE(An) and =ISERROR(An) beside each
-- cell, calculate, read their string values, and close without saving. Returns one line per row:
-- shown, class, ERROR.TYPE, ISERROR, tab-separated. Non-interactive: alerts are off while it runs
-- and restored on every exit path. Touches no other open workbook.
on run argv
	set xlPath to item 1 of argv
	set xlCopy to item 2 of argv
	set n to (item 3 of argv) as integer
	set xlName to do shell script "basename " & quoted form of xlPath
	set xlCopyName to do shell script "basename " & quoted form of xlCopy
	set tb to tab -- the separator, taken before the tell block
	set lf to linefeed
	set out to ""
	tell application "Microsoft Excel"
		if exists workbook xlName then error "already open in Excel, close it first: " & xlName
		if exists workbook xlCopyName then error "already open in Excel, close it first: " & xlCopyName
		set display alerts to false
		try
			open workbook workbook file name xlPath
			repeat 20 times
				if exists workbook xlName then exit repeat
				delay 0.25
			end repeat
			set xlSheet to worksheet 1 of workbook xlName
			set shown to {}
			repeat with r from 1 to n
				set xlCell to range ("A" & r) of xlSheet
				set t to string value of xlCell
				try
					set k to (class of (value of xlCell)) as text
				on error
					set k to "?"
				end try
				set end of shown to {t, k}
			end repeat
			save workbook as (workbook xlName) filename xlCopy file format workbook normal file format
			set xlSheet to worksheet 1 of workbook xlCopyName
			repeat with r from 1 to n
				set formula of range ("B" & r) of xlSheet to "=ERROR.TYPE(A" & r & ")"
				set formula of range ("C" & r) of xlSheet to "=ISERROR(A" & r & ")"
			end repeat
			calculate
			repeat with r from 1 to n
				set {t, k} to item r of shown
				set e to string value of range ("B" & r) of xlSheet
				set i to string value of range ("C" & r) of xlSheet
				set out to out & t & tb & k & tb & e & tb & i & lf
			end repeat
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
	return out
end run

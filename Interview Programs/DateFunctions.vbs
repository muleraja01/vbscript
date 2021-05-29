Rem DateAdd  =  Returns a date to which a specified time interval has been added
Rem DateDiff =	Returns the number of intervals between two dates
Rem IsDate   = Returns a Boolean value that indicates if the evaluated expression can be converted to a date

Rem DateDiff 
'Syntax : DateDiff(interval, date1, date2, firstDayOfWeek, firstWeekOfYear)
' q - Quarter Difference
' d - Days Difference
' w - Weeks Difference
' m - Months Difference
' h - Hours Difference
' n - Minutes Difference
' s - Seconds Diffference

Dim Date1, Date2, x

Date1=#10-03-21#
Date2=#21-10-21#

q=DateDiff("q", Date1, Date2)
d=DateDiff("d", Date1, Date2)
w=DateDiff("w", Date1, Date2)
m=DateDiff("m", Date1, Date2)

MsgBox q
MsgBox d
MsgBox w
MsgBox m
dt=Date()
dN=Now()

MsgBox "Date :" &dt
MsgBox "Date And Time :" & dN

MsgBox "Current Day : " & Day(dt)
MsgBox "Current Month : " & Month(dt)
MsgBox "Current Month Full Name : " & MonthName(Month(dt),True)
MsgBox "Current Yesr : " & Year(dt)
MsgBox "Current HOur : " & Hour(dN)
MsgBox "Current Minute : " &  Minute(dN)
MsgBox "Current Second : " & Second(dN)

'Interval, a Required Parameter. It can take the following values −
'd − day of the year.
'm − month of the year
'y − year of the year
'yyyy − year
'w − weekday
'ww − week
'q − quarter
'h − hour
'm − n
's − second

MsgBox DatePart("yyyy", dN)
MsgBox "Week Of Year : " & DatePart("ww", dN)
MsgBox "Week Of Day  : " & DatePart("w", dN)
MsgBox "Date Serials ->Format Date" & DateSerial(Year(dt),Year(dt),Day(dt))
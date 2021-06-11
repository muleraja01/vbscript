Rem DateAdd(interval,number,date)
'yyyy - Year
'q - Quarter
'm - Month
'y - Day of year
'd - Day
'w - Weekday
'ww - Week of year
'h - Hour
'n - Minute
's - Second

Y=DateAdd("yyyy",1,"31-Jan-2021")
q=DateAdd("q",1,"31-Jan-2021")
m=DateAdd("m",1,"31-Jan-2021")
y=DateAdd("y",1,"31-Jan-2021")
d=DateAdd("d",1,"31-Jan-2021")
w=DateAdd("w",1,"31-Jan-2021")
ww=DateAdd("ww",1,"31-Jan-2021")
h=DateAdd("h",1,"31-Jan-2021")
n=DateAdd("n",1,"31-Jan-2021")
s=DateAdd("s",1,"31-Jan-2021")

MsgBox "Year:" & Y
MsgBox  "Quarter:" &q
MsgBox  "Month:" &m
MsgBox  "Day Of Year:" &y
MsgBox  "Day :" &d
MsgBox  "WeekEnd:" &w
MsgBox  "Week of Year:" &ww
MsgBox  "Hour:" & h
MsgBox  "Minute:" &n
MsgBox  "Second:" &s


Rem How to Check Given Year is Leap Year Format

leapYear=InputBox("Please enter the Year")

IF IsDate(leapYear) mod 4=0 Then
	MsgBox "The Given Year " & leapYear & " is Leap Year"
Else
	MsgBox "The Given Year " & leapYear & " is Not Leap Year"
End IF


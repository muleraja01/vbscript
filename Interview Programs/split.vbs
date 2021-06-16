Rem Split


Rem Syntax: Split(expression, delimeter, Count, compare)- Delimeter -> 	subString 
Rem													-	count Optional. The number of substrings to be returned. -1 indicates that all substrings are returned
		
		
a=Split("SundayMondayTuesdayWednesdayThursdayFridaySaturday","day",3,vbTextCompare)
'for each x in a
 '   MsgBox x
'next

For i= 0 TO UBound(a) Step 1
	MsgBox a(i)
Next 

Str="SundayMondayTuesdayWednesdayThursdayFridaySaturday"
Set reex=New RegExp
reex.pattern="day"
reex.Global=True
reex.ignoreCase=True
Set obj=reex.Execute(Str)

for Each i in obj
	'msgBox i.value
Next

str=InputBox("enter the Level")

For i= 1 To str step 1
	For j= 1 To i
		newstr=newstr & i
	Next
	newstr=newstr &vbNewLine

Next 
msgBox newstr

For i= str to 1 step -1
	For j= 1 to i
		newstr=newstr & i
	NExt
	newstr=newstr &vbNewLine
Next
msgBox newstr

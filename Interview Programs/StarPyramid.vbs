
str=InputBox("enter the Level")

For i= 1 To str step 1
	For j= 1 To i
		If j=1 Then
			newstr="*"
		Else
			newstr=newstr & vbTab & "*"
		End If
	Next
	newstr=newstr &vbNewLine
Next 
msgBox newstr

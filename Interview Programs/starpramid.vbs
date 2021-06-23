str=InputBox("Enter The Level")

For i= 1 TO str Step 1
	For j=str To i Step -1
		newstr=" "
	Next
	For k= 1 To i step 1
		newstr="*" & " "
	Next
	'newstr=""
Next
MsgBox newstr
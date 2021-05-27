On Error Resume Next
intVar=InputBox("Please Enter the Integers")

For i= 1 To len(intVar) Step 1
		strvar=mid(intVar,i,1)
		
	Select case CInt(strvar)
		Case 1
			newStr=newStr&"One"
		Case 2
			newStr=newStr&"Two"
		Case 3
			newStr=newStr&"Three"
		Case 4
			newStr=newStr&"Four"
		Case 5
			newStr=newStr&"Five"
		Case 6
			newStr=newStr&"Six"
		Case 7
			newStr=newStr&"Seven"
		Case 8
			newStr=newStr&"Eight"
		Case 9
			newStr=newStr&"Nine"
		Case 0
			newStr=newStr&"Zero"	
		Case else
			MsgBox "Please Enter Only Integers"
	End Select
	

	
Next

MsgBox newStr
On Error Goto 0
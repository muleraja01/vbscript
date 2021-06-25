str= "Hi, How are you?"
newStr=""

For i=1 TO len(str) step 1
	IF Instr(newStr,mid(str, i, 1))=0 Then
		newStr=newStr&mid(str,i,1)
	else
		newStr=Replace(newStr, mid(str,i,1),"",1,1, vbTextCompare)
		Exit For
	End IF
	
Next
MsgBox newStr



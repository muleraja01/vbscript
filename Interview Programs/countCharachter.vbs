
Rem Count of Each Charachter
str="ABHGYHHOUIHHGYUUHU"
newStr=""

For i=1 To len(str)
	If Instr(newStr, mid(str,i,1))=0 Then
		newStr=newStr&mid(str,i,1)
		count=len(str)-len(Replace(str,mid(str, i, 1),""))
		MsgBox "Count of "& mid(str,i,1) &" is "&count
	End If
Next
		

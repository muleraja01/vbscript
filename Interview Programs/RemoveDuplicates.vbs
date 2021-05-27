Rem Remove Duplicates from a String

str="MJJFNIOFNKJHHK"
newStr=""

For i=1 To len(str)
	If Instr(newStr,mid(str, i, 1))=0 Then
		newStr=newStr & mid(str, i, 1)
	End IF
Next
MsgBox newStr
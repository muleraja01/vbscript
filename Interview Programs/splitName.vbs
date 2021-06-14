str="Raja Sekhar Reddy"

newStr=Split(str,"a",4,vbTextCompare)
MsgBox UBound(newStr)

For Each i in newStr

	MsgBox i
Next
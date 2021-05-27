str="Raja Amit Pavan Sekhar Text Pavan Pavan Raja Raja Pavan"
Set arrayList=CreateObject("System.Collections.ArrayList")
newStr=Split(str," ")

For i=0 To UBound(newStr)
	If arrayList.contains(newStr(i)) Then
		arrayList.add newStr(i)
		count=split(str, newStr(i))
		MsgBox newStr(i)&" is:"&UBound(count)		
	End IF
Next


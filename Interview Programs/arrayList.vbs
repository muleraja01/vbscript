arr=Array(4,6,9,7,4,5,6,7)
Set arrayList=CreateObject("System.Collections.ArrayList")

For i = 0 To Ubound(arr) 
	arrayList.add arr(i)
Next

arrayList.Sort()
MsgBox Join(arrayList.ToArray,VbNewline)'
arrayList.Reverse
MsgBox Join(arrayList.ToArray,VbNewline)


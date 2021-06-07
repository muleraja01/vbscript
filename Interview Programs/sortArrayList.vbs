
arr=Array(4,6,9,7,4,5,6,7)
Set arrayList=CreateObject("System.Collections.ArrayList")
'For i=0 To UBound(arr) Step 1
'	arrayList.Add arr(i)
'Next
For Each i in arr
arrayList.Add i
Next
arrayList.Sort()

MsgBox Join(arrayList.ToArray,vbNewLine)
arrayList.Reverse()
MsgBox Join(arrayList.ToArray,vbNewLine)

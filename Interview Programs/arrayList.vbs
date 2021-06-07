arr=Array(4,6,9,7,4,5,6,7)
Set arr=CreateObject("System.Collections.ArrayList")

'For i = 0 To Ubound(arr) 
'	MsgBox arr(i)
'Next

arr.Sort()
MsgBox Join (arr.ToArray, vbNewLine)


Rem Remove Duplicates from a array
Option Explicit
Dim arr, newarr,arrayList,i,j
arr=Array("Raja","Sekhar","Reddy","Raja","Rajani","Mrunal")
newarr = ""
Set arrayList=CreateObject("System.Collections.ArrayList")
For i= 0 To UBound(arr)
	If InStr(newarr,arr(i))=0 Then
		If i=0 Then
			newarr=newarr & " " & arr(i)
		Else
			newarr=newarr & " ," & arr(i)
		End IF
	End	IF
Next
MsgBox newarr
For j= 0 To UBound(arr)
	IF Not arrayList.Contains(arr(j)) Then
		arrayList.add arr(j)
	End IF
Next

MsgBox Join (arrayList.ToArray, vbNewLine)
	


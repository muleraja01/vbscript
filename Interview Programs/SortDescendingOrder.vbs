
Rem Sort the array List In descending order

Dim arr , temp, i, j

arr=Array(8,5,4,1,3,2,6)

For i=UBound(arr)-1 To 0 Step-1
	For j= 0 To i 
		If arr(j)<arr(j+1) Then
		temp=arr(j+1)
		arr(j+1)=arr(j)
		arr(j)=temp
		End If
	Next
Next
MsgBox Join(arr,vbNewLine)
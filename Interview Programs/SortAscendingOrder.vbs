Rem Sorting an array in ascendind Order
option Explicit
Dim arr, i, j,temp
arr=Array(4,6,9,7,4,5,6,7)

For i=UBound(arr)-1 To 0 Step -1
	For j=0 To i 
		if arr(j)>arr(j+1) then
			temp=arr(j+1)
			arr(j+1)=arr(j)
			arr(j)=temp
		End if
	Next
Next

MsgBox(arr)
Dim arr()
Redim arr(2)
arr(0)=1
arr(1)="Raja"
arr(2)="Sekhar"


For i=0 to uBound(arr)
	msgbox arr(i)
Next

Redim preserve arr(3)
arr(3) ="Reddy"


For i=0 to uBound(arr)
	msgbox arr(i)
Next
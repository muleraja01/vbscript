Rem HOw to Compare two bitMap Files

Dim FirstBmp, SecondBmp

FirstBmp ="C:\Users\Img.bmp"
SecondBmp="C:\Users\Img2.bmp"

Set myCompare=CreateObject("Mercury.FileCompare")

If myCompare.IsEqualBin(SecondBmp,FirstBmp,0,1) Then
	Print Pass
Else
	Print Fail
End IF
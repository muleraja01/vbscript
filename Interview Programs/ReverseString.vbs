Rem String Reverse WithOut strReverse Fucntion
Dim strName
strName="Autoamtion"
newStr=""

For i= len(strName) To 1 Step -1
	newStr=newStr & mid(strName,i,1)
Next
MsgBox newStr

str="Autoamtion"
Set regExp=New RegExp
regExp.pattern="[A-Za-z0-9]"
regExp.Global=True
set obj=regExp.Execute(str)

For Each i in obj
	result=i.value &result
Next
MsgBox result
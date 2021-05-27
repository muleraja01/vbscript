str="UJRKJHFSDJjkln0987.*"
Set regExp=New RegExp
regExp.pattern="[A-Za-z0-9]"
regExp.Global=True
set obj=regExp.Execute(str)
set obj2=regExp.Execute(str)
Rem -> Return The Reverse String
For Each i in obj
	result=i.value & result
Next 
Msgbox result
Rem Return not String pattern
For j=0 To obj2.count-1
	newresult = newresult&obj2.item(j)
Next 
MsgBox newresult
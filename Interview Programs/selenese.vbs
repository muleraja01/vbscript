Str="Selenese"
Set reex= New RegExp
reex.Pattern ="e"
reex.Global=True
Set Obj =Reex.Execute(Str)
MsgBox  Obj.Count
For i =1 To Obj.Count
newStr=""
For j= 1 To i 
	newStr=newStr & "*"
Next
	Str=Replace(Str, "e", newStr, 1, i,vbTextCompare)
Next	

Msgbox Str

Str="Selenese"
Set reex= New RegExp
reex.Pattern ="e"
reex.Global=True
Set Obj =Reex.Execute(Str)
MsgBox  Obj.Count
For i =1 To Obj.Count
Str=Replace(Str, "e", String(i,"*"), 1, i,vbTextCompare)
Next	

Msgbox Str






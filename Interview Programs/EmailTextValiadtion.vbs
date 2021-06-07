
Str="rajasekhar965.official@gmail.com"
Set reex=New RegExp

reex.Pattern="^\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b$"
reex.IgnoreCase=True

 Obj=reex.Test(Str)

MsgBox obj
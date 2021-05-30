Str="missisippie"
count=Split(str,"i")

MsgBox Ubound(count)

Set regEx=New RegExp
regEx.Pattern = "i"
regEx.Global= True
MsgBox regEx.Test(Str)
set objReg=regEx.Execute(Str)

MsgBox objReg.count



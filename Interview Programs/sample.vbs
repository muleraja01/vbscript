Str="missisippie"
count=Split(str,"i")

MsgBox Ubound(count)-1

Set regEx=New RegExp
regEx.Pattern = "i"
regEx.Global= True

set objReg=regEx.Execute(Str)

MsgBox objReg.count
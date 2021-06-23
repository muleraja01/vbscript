str="Raja Sekhar Reddy"

MsgBox InstrRev(str, Right(str,1))
MsgBox InstrRev(str, "")

Set reex=New RegExp
reex.Pattern="[a-zA-Z0-9 ]"
reex.Global=True
Set Obj=reex.Execute(str)
MsgBox len(str)
MsgBox Obj.Count
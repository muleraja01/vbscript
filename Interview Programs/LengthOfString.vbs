Rem Get Len of String without using Len Method


str="AKHUJHD"
val=InstrRev(str,Right(str,1))

MsgBox val


str="Autoamtion"
Set regExp=New RegExp
regExp.pattern="[A-Za-z0-9]"
regExp.Global=True
set obj=regExp.Execute(str)

Msgbox obj.Count

set obj=regExp.Test(str)

MsgBox


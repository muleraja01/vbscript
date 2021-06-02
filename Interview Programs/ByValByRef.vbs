Sub begin()
Dim last_column

last_column = 234
MsgBox "Begin:" & last_column

trying_byref x:=last_column
MsgBox "byref:" & last_column

trying_byval v:=last_column
MsgBox "byval:" & last_column
End Sub

Sub trying_byref(ByRef x)
x = 111
End Sub

Sub trying_byval(ByVal v)  '<--not ByRef, that was in sub trying_byref.
v = 222
MsgBox "In Here:" & v
End Sub
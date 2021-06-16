
Call begin()
Function begin()
	Call trying_byref()
	Call trying_byval()
	call trying_byref()
End Function

Sub trying_byref(ByRef x)
MsgBox "In Here:" & x
End Sub

Sub trying_byval(ByVal v)  '<--not ByRef, that was in sub trying_byref.
MsgBox "In Here:" & v
v = 222
MsgBox "In Here:" & v
End Sub
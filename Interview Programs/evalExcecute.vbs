' Eval And Execute
'EVal  returns expression
x=100
y=101
setEqual=Eval("x=y")
MsgBox "Equality is:"&setEqual
MsgBox "value of x is:" &x

'Execute
x=100
'Here Y assign vlaue 101 to x
setEqual=Execute("x=y")
MsgBox "Equality is:"& setEqual
MsgBox "value of x is:" & x


Rem Fibonacci Series

Rem 0 1 1 2 3 5 8 13

a=0
b=1

n=InputBox("Please Enter the Number")

For i=1 TO n
	temp=a+b
	MsgBox temp
	a=b
	b=temp
Next
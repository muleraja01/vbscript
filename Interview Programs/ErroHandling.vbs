Rem  On Error Resume Next
Rem Err Object
Rem On Error GoTO 0

'On Error Resume Next  : Moves the control of the cursor to the next line of the error statement

On Error Resume Next

Dim Result

result=20/0
if result = 0 Then
MsgBox "Result is 0"
Else
MsgBox "Result is non Zero"
End IF




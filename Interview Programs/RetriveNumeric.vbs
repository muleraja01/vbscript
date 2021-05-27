Rem Retrive Numberic values from main String

Dim strName, numString

strName="kjijoij89679mjk"


For i=1 To len(strName)
	IF isNumeric(mid(strName,i,1)) then
		numString=numString & mid(strName,i,1)
	End IF
Next

MsgBox numString
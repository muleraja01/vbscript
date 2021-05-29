
Rem InStr 
'Return Position of 1st Occurenece of 1 string with another String .Search begins from 1st charachter
'Syntax: Instr(Start, String1, String 2, Compare)
'Start -Option Position of charachter
'Compare :  0- VbBinaryCompare-Binary Comprassion(Default)
'			1- VbTextCompare -Text Comprassion
'String1 is  "" returns 0
'String2 is  "" returns Start
'string1 or string 2 is null ->Instr Return null
'String 2 is not found return 0

str1="My name is Raja Sekahr Reddy"

nullStr=Null;
MsgBox Instr(1, str1, "a")>0

MsgBox Instr(1, "","a")>0

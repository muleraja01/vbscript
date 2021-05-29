'===========================InStrRev==================================================
'	Return the postition of 1st occurence of the String from the last charachter .
'	Search begins at last charachter and count from the 1st charachter
'	InStrRev(String1, String2, start, compare)
'	Start -Option Position of charachter
'	Compare :   0- VbBinaryCompare-Binary Comprassion(Default)
'				1- VbTextCompare -Text Comprassion
'	String1 is  "" returns 0
'	String2 is  "" returns Start
'	string1 or string 2 is null ->Instr Return null
'	String 2 is not found return 0
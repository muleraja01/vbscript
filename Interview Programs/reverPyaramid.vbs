For i = 5 to 0 Step -1
	For j= i-1 To 0 Step 1
		str = str & String(j, "i") &vbNewLine 
		For k= i-j To 0 Step 1
			str=str & String(k, "*") &vbNewLine 
		Next
	Next
	
 
Next

MsgBox str

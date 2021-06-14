str="raja sekhar reddy"

For i= 1 to len(str)
	count=len(str)-len(replace(str, mid(str, i, 1),""))
	msgbox "The charachter for :" & mid(str,i, 1) &" is appreared: "& count
Next
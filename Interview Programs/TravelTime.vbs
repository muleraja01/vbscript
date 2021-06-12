'    [10, 30, 35, 60, 80]
travelTime=InputBox("enter the number")
arrTime=array(5, 10, 45, 20)
Set suggTime=CreateObject("System.Collections.ArrayList")
For i=0 To Ubound(arrTime)-1 step 1
	if Instr(travelTime ,arrTime(i))<=0 Then
		suggTime.add arrTime(i)
	End IF
Next

MsgBox Join(suggTime.ToArray,vbNewLine)

'MSgbox "The Travel Time is : "& travelTime & " The suggested playback  is "& suggTime.ToArray


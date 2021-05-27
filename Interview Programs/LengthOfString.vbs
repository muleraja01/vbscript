Rem Get Len of String without using Len Method


str="AKHUJHD"
val=InstrRev(str,Right(str,1))

MsgBox val


'2 Option

newstr="BAHIUDTE"
i=0

Do Until newstr=""

	i=i+1
	newstr=Replace(newstr, Left(newstr,1),"")
Loop
MsgBox i


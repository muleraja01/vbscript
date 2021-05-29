

Rem Filter(inputstrings,value[,include[,compare]])

arr=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
newStr=Filter(arr,"S")
for each i in newStr
    MsgBox i
next

newStr2=Filter(arr,"S",false)
for each j in newStr2
   MsgBox j
next

Rem VarType

x=200

y=VarType(x)
'Retrun 2 -if VarType is Integer
MsgBox "VarType" & y 

x=200.56

y=VarType(x)
'Retrun 5 -if VarType is double
MsgBox "VarType" & y 

x="Delhi"

y=VarType(x)

'8 Return 8 VarType is String
MsgBox "VarType" & y


Rem Asc

num=Asc("64")
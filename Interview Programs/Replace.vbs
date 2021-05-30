

Rem Replace

'Syntax  : Replace(string, findString, ReplaceString , start, count, Compare)


str="Raja Sekhar Reddy"

newStr=Replace(str, "a","$",1,1,vbTextCompare)
MsgBox newStr

str="Raja Sekhar Reddy"

newStr=Replace(str, "d","$",5,-1,vbTextCompare)
MsgBox newStr

str="Raja Sekhar Reddy"
newStr=Replace(str, "a","$",1,-1,vbTextCompare)
MsgBox newStr

str="Raja Sekhar Reddy"
newStr=Replace(str, "a","$",1,-1,vbBinaryCompare)
MsgBox newStr

str="Raja Sekhar Reddy"
newStr=Replace(str, "a","$",vbTextCompare)
MsgBox newStr
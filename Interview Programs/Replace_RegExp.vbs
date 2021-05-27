Rem Replace the string without using Replace Method
Rem Can be achieved using RegExp

Dim RegX

Dim myString, searchString, ReplaceString, newString
myString =" Raja has Experiece in UFT"
searchString ="UFT"
ReplaceString ="Selenium"
Set RegX=New RegExp
RegX.Pattern=searchString
RegX.Global=True
newString=RegX.Replace(myString,ReplaceString)
MsgBox newString



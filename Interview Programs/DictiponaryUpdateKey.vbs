Dim d
Set d=CreateObject("Scripting.Dictionary")
d.Add "re","Red"
d.Add "gr","Green"
d.Add "bl","Blue"
d.Add "pi","Pink"
MsgBox d.Item("re")
d.Key("re")="r"
MsgBox "The value of key r is: " & d.Item("r")
Set autoamtion = CreateObject("Scripting.Dictionary")
autoamtion.Add "Raja", "UFT"
autoamtion.Add "Pavan", "Selenium"
autoamtion.Add "Amit", "Pythn"

MsgBox autoamtion.Item("Raja")
Set skillSet = CreateObject("Scripting.Dictionary")
skillSet.Add autoamtion.Item("Raja"), "VbScript"
skillSet.Add autoamtion.Item("Pavan"), "Java"
Msgbox skillSet.Item("UFT")
Msgbox skillSet.Item("Selenium")
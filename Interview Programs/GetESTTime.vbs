strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_TimeZone")
For Each objItem in colItems
    Set colItems2 = objWMIService.ExecQuery("Select * From Win32_ComputerSystem")
    For Each objItem2 in colItems2
        blnDaylightInEffect = objItem2.DaylightInEffect
    Next
    If blnDaylightInEffect Then
        Wscript.Echo objItem.DaylightName
    Else
        Wscript.Echo objItem.StandardName
    End If
Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
	WScript.Echo "Current Time Zone (Hours Offset From GMT): " & (objItem.CurrentTimeZone / 60)
	WScript.Echo "Daylight Saving In Effect: " & objItem.DaylightInEffect
Next
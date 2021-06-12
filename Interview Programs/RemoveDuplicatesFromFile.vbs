Set ObjFSo=CreateObject("Scripting.FileSystemObject")
Set ObjFile=ObjFSo.OpenTextFile("C:\Users\PC\Desktop\vbscript\Interview Programs\textFile.txt")
Set objDictionary=CreateObject("Scripting.Dictionary")


Do Until ObjFile.AtEndOfStream
strName = ObjFile.ReadLine
	If Not objDictionary.Exists(strName) Then
		objDictionary.Add strName, strName
		MsgBox objDictionary.Item(strName)
	End If
Loop
ObjFile.Close





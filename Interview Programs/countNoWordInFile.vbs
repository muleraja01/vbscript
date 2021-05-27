Const ForReading = 1
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objTextFile=objFSO.OpenTextFile("E:\Traning\Synechron_RestApi\TopicsNotes.txt", ForReading)
wordCount=0
DO Until objTextFile.AtEndOfStream
	line=objTextFile.ReadLine()
	If line<>"" Then
		str1=Split(line," ")
		wordCount=wordCount+ubound(str1)+1
	End If
Loop
	MsgBox wordCount
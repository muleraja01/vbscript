Const ForReading = 1
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objTextFile=objFSO.OpenTextFile("E:\Traning\Synechron_RestApi\TopicsNotes.txt", ForReading)

'Read the count of lines
'objTextFile.ReadAll
'strlineCount=objTextFile.Line
'MsgBox strlineCount
' To Read Each 20 Lines
'StrNumberTimes=Round((strlineCount/20),0)
'MsgBox StrNumberTimes
i=0
DO Until objTextFile.AtEndOfStream
	line=objTextFile.ReadLine()
	If line<>"" Then
		fileds=Split(line," ")
		'MsgBox fileds
		i=i+1
	End If
Loop
	MsgBox i
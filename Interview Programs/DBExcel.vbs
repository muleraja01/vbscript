Dim objConn,objRs,strExcelFile,strQuery

Set objConn=CreatObject("ADODB.Connection")
Set objRs=CreateObject("ADODDB.RecordSet")
strExcelFile=""
objConn.Open "Driver={Microsoft Excel Driver (*.xls)};DBQ= "&strExcelFile& ";ReadOnly=True;
strQuery="Select * From [TestData$]"
objRs.Open objConn , strQuery
value=objRs.fields.item(0)
MsgBox value

objRs.close
objConn.close
Set objRs=Nothing
Set objConn=Nothing
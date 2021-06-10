QCUtil Object 
QCUtil Object is used to access QC OTA.
Below are the properties of QCUtil object that can be used for extracting information from QC.

Properties of QCUtil Object

1. QCUtil.CurrentRun:  This give reference to the current test run in QC.We can then get  information of currentrun  like name as shown below:

set objRun = QCUtil.CurrentRun
msgbox objRun.Name

2. QCUtil.CurrentTest: This give reference to the current test in QC. We can extract information like name of the test, adding attachment to current test using this property

set objTest = QCUtil.CurrentRun
msgbox objTest.Name

3. QCUtil.CurrentTestSet: This give reference to the current test set in QC. We can  extract information similar as above for testset using this property.

set CurrentTSTest = QCUtil.CurrentTestSet

4. QCUtil.CurrentTestSetTest: This gives reference to current test execution instance in the testset.

set CurrentTSTestSet = QCUtil.CurrentTestSetTest

5. IsConnected Property: returns true/false based on QTP currently connected to QC project.

blnQCsts = QCUtil.IsConnected

6. QCConnection Property: This gives an instance of current run session and can access the structure of QC using this property

Set QCConnection = QCUtil.QCConnection


Examples of working with QCUtil Object

1.Adding a defect in QC using QCUtil Object:

            ‘Create instance of QCConnection
            Set QCConnection = QCUtil.QCConnection
‘Create an instance of BugFactory
Set DefFactory = QCConnection.BugFactory
'Add a new defect
Set Bug = DefFactory.AddItem(Nothing)
‘Provide mandatory details for the defect
Bug.Status = “New”
Bug.Summary = “Module Detected new defect summary”
Bug.DetectedBy = “njoshi”
Bug.AssignedTo = “dev001”
Bug.Post
Set DefFactory = nothing
Set QCConnection = nothing.

2. Adding Attachment to QC 

Dim ObjCurrentTest,ObjAttch
‘ We can add attachment to currentTest, current run, testset, and testsettest by using repective ‘properties 
Set ObjTest = QCUtil.CurrentTest.Attachments
Set ObjAttachFile = ObjTest.AddItem(Null)
ObjAttachFile.FileName = FileName ‘ Provide path of the file that needs to be attached to QC 
ObjAttachFile.Type = 1
ObjAttachFile.Post
ObjAttachFile.Refresh

3. Validating If QC is connected properly: 

if QCUtil.IsConnected then
msgbox “QC is connected”
Else
Msgbox “QC is not connected”
EndIf
   
4. Connecting to QC through QTP

Set qtApp = CreateObject ("QuickTest.Application")
If  qtApp.launched <> True then
     qtApp.Launch
End If
qtApp.Visible = "true"
If Not qtApp.TDConnection.IsConnected Then
      qtApp.TDConnection.Connect QCurl, DomainName, ProjectName, UserName, Password, False
End If


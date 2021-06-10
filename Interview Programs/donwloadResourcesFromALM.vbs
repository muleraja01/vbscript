fn_QCDownloadResource "Config.xml","C:TempTest"
'*******************************************Sub***************************************************
'Function Name:-        fn_QCDownloadResource
'Function Description:- Search in "Test Resources" tab and download it 
'Input Parameters:-     sFileName(File to Download),sDestination(Optional, Folder to download the File)
'Output Parameters:-    None
'**************************************************************************************************
Public Function fn_QCDownloadResource(sFileName, sDestination)

 Dim objQC, objRes, objFilter, objFileList, objFile, sFilterPair, iCount
 'Check the ALM + UFT connection
 If Not QCUtil.IsConnected Then
  Msgbox "Please Connect ALM with UFT"
  Exit Function
 End If 
 'Creating ALM OTA QCConnection object,for accessing the ALM object model. 
 Set objQC = QCUtil.QCConnection
 'Assessing the Resources inside the Test Resources
 Set objRes = objQC.QCResourceFactory
 'Activating the Filter object
 Set objFilter =  objRes.Filter
 'Matching the given FileName in the Test Resources files
 objFilter.Filter("RSC_FILE_NAME") = """" & Cstr(sFileName) & """"
 'Creates a list of objects according to the specified filter.
 Set objFileList = objFilter.NewList
 'Incase Download path is empty, file is downloaded into Temp Folder
 If sDestination = "" Then
   sDestination = "C:Temp"
 End if 
 'Downloading the found File
 If objFileList.Count = 1 Then
   Set objFile = objFileList.Item(1)
   objFile.FileName = sFileName
   'Downlaod the file and sync= True , for completion of process
   objFile.DownloadResource sDestination, True
   Msgbox "'"&sFileName&"',Successfully downloaded to '"&sDestination&"'"
 Else
   Msgbox "Sorry No Match found in ALM for the FileName = '"&sFileName&"'"
 End If
 'Destroying the Objects  
 Set objQC = Nothing
 Set objRes = Nothing
 Set objFilter = Nothing
 Set objFileList = Nothing
 Set objFile = Nothing
End Function
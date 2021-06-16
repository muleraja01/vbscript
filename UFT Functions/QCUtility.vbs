Set QCConnection=QCUtil.Connection
Set TreeManager=QCConnection.TreeManager

strLocalFilePath=""
QCPath="Subject/Autoamtion/Set"

Set node=TreeManager.nodeByPath(QCPath)
Set attachments=node.attachments

Set att=attachments.AddItem(Null)
att.FileName=strLocalFilePath
att.Post()
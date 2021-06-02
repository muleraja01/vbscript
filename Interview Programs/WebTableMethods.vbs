Rem WebTable Methods.

Obj.RowCount
Obj.ColCount(Row)

'GetCellData
Obj.GetCellData(Row,Column)

'GetRowWithCellText
Obj.GetRowWithCellText(Text, Column, StartFromRow)'Column -Optional	'StartFromRow-Optional
						
'ChildItem
Obj.ChildItem(Row, Column, micClass, Index) 'Obj.ChildITem(2, 1, "WebCheckBox", 1)

'ChildItemCount
Obj.ChildItemCount(row, column, micClass)

'ChildObject
Set ODesc=Description.Create

oDesc("micClass").value="WebCheckBox"

Set Obj=Browser("name:=a").page("title:=abc").WebTable("name:=table).ChildObject(oDesc)

For i=0 To Obj.count-1
	obj(i).GetRoProperty("Name")
Next

Attribute VB_Name = "DupSheet"
Sub Duplicatesheet()
Attribute Duplicatesheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Duplicatesheet Macro
'

'
  ActiveSheet.Select
  ActiveSheet.Copy Before:=Sheets(5)
  ActiveSheet.Shapes.Range(Array("Button 1")).Select
    Selection.Delete
End Sub



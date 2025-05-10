Attribute VB_Name = "Module1"
Sub RefreshPivotTables()
Attribute RefreshPivotTables.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RefreshPivotTables Macro
'

'
    Range("B6").Select
    ActiveWorkbook.RefreshAll
    Range("G5").Select
    ActiveWorkbook.RefreshAll
    Range("L4").Select
    ActiveWorkbook.RefreshAll
    Sheets("Sales Dashboard").Select
    
' Navigate to Sales Dashboard sheet
Sheets("Sales Dashboard").Activate

' Show confirmation message
MsgBox "Dashboard refreshed successfully!", vbInformation, "Done"

End Sub

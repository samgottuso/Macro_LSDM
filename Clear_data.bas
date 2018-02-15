Attribute VB_Name = "Module2"
Sub clear_data()

Application.ScreenUpdating = False


Worksheets("Event.Data").Activate

'Last Row of our event data'
last_row_Event = Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

Set cleared_range = ActiveSheet.Range("A2:I" & last_row_Event)

If last_row_Event > 1 Then
    cleared_range.ClearContents
Else
    MsgBox "Data is already Cleared"
End If




End Sub

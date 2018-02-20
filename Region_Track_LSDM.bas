Attribute VB_Name = "Module3"
Sub track_region()

Application.ScreenUpdating = False


Worksheets("Event.Data").Activate

'Last Row of our event data'
last_row_Event = Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

'Roster Last Row'
Last_Row_Roster = Sheets("Crosscheck").Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row


Set new_ID_Range = ActiveSheet.Range("I2:I" & last_row_Event)

Set new_Track_Range = ActiveSheet.Range("L1:L" & last_row_Event)

Set new_Region_Range = ActiveSheet.Range("K1:K" & last_row_Event)

Set new_Org_Range = ActiveSheet.Range("G1:G" & last_row_Event)
Set Org_Range = ActiveSheet.Range("F1:F" & last_row_Event)




For Each cell In new_ID_Range:
    
    'First Enusre that no "Not Found Values" Remain
    If InStr(1, cell.Value, "Not Found") > 0 Then
        MsgBox "Please ensure that all IDs have been cleaned"
        Exit Sub
    End If
Next


    
    'If we pass that check then start assigining Tracks and Regions (as long as it's a not available)'
For Each cell In new_ID_Range:
    ID = ActiveSheet.Cells(cell.Row, 9).Value
    Region_abrv = (Application.Index(Sheets("Crosscheck").Range("C2:C" & Last_Row_Roster), Application.Match(ID, Sheets("Crosscheck").Range("A2:A" & Last_Row_Roster), 0), 1))
    'Re-run cleaned IDs for correct Org names'
    Org_check = (Application.Index(Sheets("Crosscheck").Range("B2:B" & Last_Row_Roster), Application.Match(ID, Sheets("Crosscheck").Range("A2:A" & Last_Row_Roster), 0), 1))
    
    If ID = "Not Available" Then
        new_Org_Range.Cells(cell.Row).Value = Org_Range.Cells(cell.Row).Value
    Else
        new_Org_Range.Cells(cell.Row).Value = Org_check
        
    End If
    
    
    
    If ID <> "Not Available" Then
        new_Region_Range.Cells(cell.Row).Value = Region_abrv
        new_Track_Range.Cells(cell.Row).Value = Left(ID, 2)
    Else
        new_Region_Range.Cells(cell.Row).Value = "Not Found"
        new_Track_Range.Cells(cell.Row).Value = "Not Available"
    End If
    
Next

End Sub

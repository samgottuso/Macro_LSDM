Attribute VB_Name = "Module1"

Sub auto_LSDM()

Application.ScreenUpdating = False


Worksheets("Event.Data").Activate


'Set up/Clear Columns'


For i = 1 To 20
    lrow = Cells(Rows.Count, i).End(xlUp).Row
    If InStr(1, Columns(i).Cells(1).Value, "New.") > 0 Then
        ActiveSheet.Range(Cells(2, i), Cells(lrow, i)).ClearContents
    End If
Next i

    

    

'Last Row of our event data'
last_row_Event = Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row



'Figure out our last ranges on the Crosscheck sheet'

last_row_LSDM = Sheets("Crosscheck").Cells(ActiveSheet.Rows.Count, "F").End(xlUp).Row

last_row_Distro = Sheets("Crosscheck").Cells(ActiveSheet.Rows.Count, "M").End(xlUp).Row

last_row_Roster = Sheets("Crosscheck").Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row



'Find New ID through the LSDM/Distro List '


'Put our Data into the new Columns and then perform index matches in place'

Set Email_range = ActiveSheet.Range("A2:A" & last_row_Event)

Set ID_Range = ActiveSheet.Range("H2:H" & last_row_Event)

Set new_ID_Range = ActiveSheet.Range("I1:I" & last_row_Event)

Set new_Org_Range = ActiveSheet.Range("G1:G" & last_row_Event)

Set Org_Range = ActiveSheet.Range("F1:F" & last_row_Event)

Set new_Role_Range = ActiveSheet.Range("C1:C" & last_row_Event)

Set new_Role_Other_Range = ActiveSheet.Range("E1:E" & last_row_Event)



'Index/Matches'

For Each Cell In ID_Range:
    Email = ActiveSheet.Cells(Cell.Row, 1).Value
    ID_check = (Application.Index(Sheets("Crosscheck").Range("J2:J" & last_row_LSDM), Application.Match(Email, Sheets("Crosscheck").Range("F2:F" & last_row_LSDM), 0), 1))
    ID_Check_2 = (Application.Index(Sheets("Crosscheck").Range("N2:N" & last_row_Distro), Application.Match(Email, Sheets("Crosscheck").Range("M2:M" & last_row_Distro), 0), 1))
    
    'ID Check'
    If Not IsError(ID_check) Then
        new_ID_Range.Cells(Cell.Row).Value = ID_check
    ElseIf Not IsError(ID_Check_2) Then
        new_ID_Range.Cells(Cell.Row).Value = ID_Check_2
    Else
        new_ID_Range.Cells(Cell.Row).Value = "Not Found"
    End If
    'Org Check'
    ID = new_ID_Range.Cells(Cell.Row).Value
    Org_check = (Application.Index(Sheets("Crosscheck").Range("B2:B" & last_row_Roster), Application.Match(ID, Sheets("Crosscheck").Range("A2:A" & last_row_Roster), 0), 1))
    
    If ID = "Not Found" Then
        new_Org_Range.Cells(Cell.Row).Value = "NA"
    ElseIf ID = "Not Available" Then
        new_Org_Range.Cells(Cell.Row).Value = Org_Range.Cells(Cell.Row).Value
    Else
        new_Org_Range.Cells(Cell.Row).Value = Org_check
        
    End If
    
    'Role Check'
    Role_Check = (Application.Index(Sheets("Crosscheck").Range("G2:G" & last_row_LSDM), Application.Match(Email, Sheets("Crosscheck").Range("F2:F" & last_row_LSDM), 0), 1))
    
    If Not IsError(Role_Check) Then
        new_Role_Range.Cells(Cell.Row).Value = Role_Check
    Else
        new_Role_Range.Cells(Cell.Row).Value = "Not Found"
    End If
    
    'Role Other'
    Role = new_Role_Range.Cells(Cell.Row).Value
    Role_Other_Check = (Application.Index(Sheets("Crosscheck").Range("H2:H" & last_row_LSDM), Application.Match(Email, Sheets("Crosscheck").Range("F2:F" & last_row_LSDM), 0), 1))
    
    If Role = "Other" Then
        new_Role_Other_Range.Cells(Cell.Row).Value = Role_Other_Check
    End If
    
        
    
       
    
    
    
    
    
    
    
Next














End Sub

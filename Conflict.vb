Private Sub Workbook_Open()

Dim top_row_num As Double
Dim top_row As Range
Dim rng As Range
Dim temp_top_row_num As Double
Dim temp_top_row
Dim cell As Range
Dim counter As Integer
Dim data_rng As Range

Dim data_sht As Worksheet
Dim wb As Workbook
Dim temp_workbook As Workbook
Dim temp_sheet As Worksheet
Dim bottom_row As Integer
Dim form_rng As Range
Dim form_rng_start As Range
Dim form_rng_end As Range
Dim consult_list As Range
Dim happened()
Dim new_wb As Workbook
Dim use_permission As Office.UserPermission
Dim start_date_str As String
Dim end_date_str As String
Dim start_date As Date
Dim end_date As Date
Dim new_dates As String

Dim out_app As Object
Dim out_mail As Object

Dim row_count As Integer
Dim form_top_row As Range
Dim form_top_row_num As Integer


Set out_app = CreateObject("Outlook.Application")
Set out_mail = out_app.CreateItem(0)

Set wb = ActiveWorkbook
Set data_sht = ActiveSheet


If Sheets.Count = 1 Then
    
    If MsgBox("Would you like to set a custom date range for completion date other than " & Date - 13 & " - " & Date & "?", vbYesNo, "Custom Date Range") = vbYes Then
        start_date_str = InputBox("Please enter the start date for completion date (MM/DD/YY).", "Start Date")
        If IsDate(start_date_str) Then
            start_date = DateValue(start_date_str)
        Else
            MsgBox "Invalid Date"
        End If
        end_date_str = InputBox("Please ender the end date for completion date (MM/DD/YY).", "End Date")
        If IsDate(end_date_str) Then
            end_date = DateValue(end_date_str)
        Else
            MsgBox "Invalid Date"
        End If
        new_dates = "New Dates"
    End If
  
        
    
    
    data_sht.Name = "Raw Data"
    
    cells.UnMerge
    
    top_row_num = Columns(1).Find("Team Member ID", LookAt:=xlWhole).row
    Set top_row = Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight))
    
    Set rng = Range(cells(top_row_num, 1), Range(cells(top_row_num, 1).End(xlDown), cells(top_row_num, 1).End(xlToRight)))
    
    
    rng.SpecialCells(xlCellTypeLastCell).Offset(0, 1).Copy
    rng.PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd
    
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Workbooks.Open _
            "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Reference Files - DO NOT CHANGE\Conflicting Title and Comp Change Audit - Template.xlsx"
    Application.DisplayAlerts = True
    
    Set temp_workbook = ActiveWorkbook
    Set temp_sheet = temp_workbook.Sheets("Data")
    temp_top_row_num = Columns(1).Find("Team Member ID", LookAt:=xlWhole).row
    Set temp_top_row = Range(cells(temp_top_row_num, 1), cells(temp_top_row_num, 1).End(xlToRight))
    
    
    counter = 0
    
    For Each cell In top_row
        If cell.Value <> temp_sheet.cells(temp_top_row_num, cell.Column).Value And counter = 0 Then
            MsgBox "Please Provide Correct Data Set"
            counter = counter + 1
        End If
    Next
    
    wb.Activate
    
    Workbooks("Conflicting Title and Comp Change Audit - Template.xlsx").Sheets(Array("Data", "References")).Move After:=wb.Sheets(Sheets.Count)
    
    Set temp_sheet = Sheets("Data")
    Set temp_top_row = temp_sheet.Range(cells(temp_top_row_num, 1), cells(temp_top_row_num, 1).End(xlToRight))
    
    data_sht.Activate
    
    Set data_rng = Range(cells(top_row_num + 1, 1), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column))
    data_rng.Copy
    
    temp_sheet.Activate
    cells(temp_top_row_num + 1, 1).PasteSpecial xlPasteValues
    Set form_rng_start = cells(temp_top_row_num - 1, 1).End(xlToRight)
    Set form_rng_end = cells(temp_top_row_num - 1, 1).End(xlToRight).End(xlToRight)
    
    
    
    Set form_rng = Range(form_rng_start, form_rng_end)
    form_rng.Copy
    
    Range(form_rng_start.Offset(2, 0), cells(cells(temp_top_row_num, 1).End(xlDown).row, form_rng_end.Column)).PasteSpecial xlPasteFormulas
    
    If new_dates = "New Dates" Then
        Rows(temp_top_row_num).Find("In Last Two Weeks?").Offset(-3, 0).Value = start_date
        Rows(temp_top_row_num).Find("In Last Two Weeks?").Offset(-2, 0).Value = end_date
    End If
    
    Application.Calculate
    
    Range(form_rng_start.Offset(2, 0), cells(cells(temp_top_row_num, 1).End(xlDown).row, form_rng_end.Column)).Copy
    Range(form_rng_start.Offset(2, 0), cells(cells(temp_top_row_num, 1).End(xlDown).row, form_rng_end.Column)).PasteSpecial xlPasteValues
    
    happened = Array("Both Changed", "Nothing Changed")
    
    filter_and_delete happened, temp_top_row, temp_top_row_num, temp_top_row.Find("What Happened?").Column
    filter_and_delete "Yes", temp_top_row, temp_top_row_num, temp_top_row.Find("RHI Leadership").Column
    filter_and_delete "Yes", temp_top_row, temp_top_row_num, temp_top_row.Find("Zupteam").Column
    filter_and_delete "No", temp_top_row, temp_top_row_num, temp_top_row.Find("In Last Two Weeks?").Column
    filter_and_delete "Yes", temp_top_row, temp_top_row_num, temp_top_row.Find("Record on Master List?").Column
    
    row_count = Range("A5", Range("A5").End(xlDown)).Column.Count - 1
    
    Sheets("Formatted").Activate
    Range("A13", Range("A13").End(xlToRight)).Copy
    Range("A14", Range("A13").Offset(row_count, 0)).PasteSpecial xlPasteFormulas
    Appliation.Calculate
    cells.Copy
    cells.PasteSpecial xlPasteValues
    Set form_top_row = Range("A12", Range("A12").End(xlToRight))
    form_top_row_num = 12
           
    
    Sheets("References").Activate
    Set consult_list = Range("H5", Range("H5").End(xlDown))
    
    
    For Each cell In consult_list
        Sheets("Formatted").Activate
        If cell.Value = "Jim Ziraldo" Then
            form_top_row.AutoFilter Field:=temp_top_row.Find("Consultant").Column, Criteria1:=cell.Value
            add_consult_wb form_top_row_num, usr_permissions, cell
        Else
            form_top_row.AutoFilter Field:=form_top_row.Find("Consultant").Column, Criteria1:=cell.Value
            add_consult_wb form_top_row_num, usr_permissions, cell
        End If
        form_top_row.AutoFilter Field:=form_top_row.Find("Consultant").Column, Criteria1:=cell.Value
        Sheets("Formatted").Activate
        ActiveSheet.AutoFilterMode = False
    Next
    If IsEmpty(cells(temp_top_row_num, 1).End(xlDown)) = False Then
        Workbooks.Open _
            "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Reference Files - DO NOT CHANGE\Conflicting Master List.xlsx"
        Set new_wb = ActiveWorkbook
        wb.Activate
        Sheets("Data").Range(cells(temp_top_row_num, 1).Offset(1, 0), cells(cells(temp_top_row_num, 1).End(xlDown).row, cells(temp_top_row_num, 1).End(xlToRight).Offset(0, -1).Column)).SpecialCells(xlCellTypeVisible).Copy
        new_wb.Sheets("Master").Activate
        Range("A2").End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
        Range("A2").End(xlToRight).Offset(1, 0).Copy
        Range(Range("A2").End(xlToRight).End(xlDown).Offset(1, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).PasteSpecial xlPasteFormulas
        new_wb.Save
        new_wb.Close
    End If
    
        On Error Resume Next
    With out_mail
        .To = "cariesanford@quickenloans.com; kirkwaters@quickenloans.com"
        .CC = ""
        .BCC = ""
        .Subject = "Conflicting Title/Pay Change Macro has COMPLETED!!"
        .body = ""
        .send
    End With
    On Error GoTo 0
    
    Set out_mail = Nothing
    Set out_app = Nothing
    
        
    wb.SaveAs "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Conflicting Title and Pay Change Report\Conflicting Audit - " & Month(Date) & "." & Day(Date) & "." & Year(Date) & ".xlsm"
    wb.Close
    
End If

End Sub




Function filter_and_delete(the_array, temp_top_row, temp_top_row_num, column_num)

If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If

temp_top_row.AutoFilter Field:=column_num, Criteria1:=the_array, Operator:=xlFilterValues

If IsEmpty(cells(temp_top_row_num, 1).Offset(1, 0)) = False And IsEmpty(cells(temp_top_row_num, 1).Offset(2, 0)) = False Then
    Range(cells(temp_top_row_num, 1).Offset(1, 0), cells(temp_top_row_num, 1).End(xlDown)).SpecialCells(xlCellTypeVisible).EntireRow.Delete Shift:=xlUp
ElseIf IsEmpty(cells(temp_top_row_num, 1).Offset(1, 0)) = False And IsEmpty(cells(temp_top_row_num, 1).Offset(2, 0)) Then
    cells(temp_top_row_num, 1).Offset(1, 0).EntireRow.Delete Shift:=xlUp
End If

ActiveSheet.AutoFilterMode = False

End Function


Function add_consult_wb(temp_top_row_num, usr_permissions, cell)

Dim new_wb As Workbook

If IsEmpty(cells(temp_top_row_num, 1).Offset(1, 0)) = False Then
    Range(cells(temp_top_row_num, 1), cells(cells(temp_top_row_num, 1).End(xlDown).row, cells(temp_top_row_num, 1).End(xlToRight).Column)).SpecialCells(xlCellTypeVisible).Copy
    Workbooks.Add
    Set new_wb = ActiveWorkbook
    Range("A2").PasteSpecial xlPasteValues
    If cell.Value = "Jim Ziraldo" Then
        Set usr_permissions = ActiveWorkbook.Permission.Add("CompensationStrategyTeam@quickenloans.com", MsoPermission.msoPermissionFullControl)
        Set usr_permissions = ActiveWorkbook.Permission.Add("JimZiraldo@quickenloans.com", MsoPermission.msoPermissionFullControl)
        new_wb.SaveAs "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Conflicting Title and Pay Change Report\" & _
            cell.Value & " - " & Month(Date) & "." & Day(Date) & "." & Year(Date) & ".xlsx"
    Else
        new_wb.SaveAs "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Conflicting Title and Pay Change Report\" & _
            cell.Value & " - " & Month(Date) & "." & Day(Date) & "." & Year(Date) & ".xlsx"
    End If
    new_wb.Close
End If


End Function

Sub CheckSheetsAndRun()

'THIS IS THE MACRO THAT KICKS IT ALL OFF

Dim start_time As Double
Dim elapsed_minutes As Double
Dim tm_count As Integer

Dim active_workbook As Workbook
Dim opened_workbook As Workbook
Dim sht As Worksheet
Dim counter As Integer
Dim counter2 As Integer
Dim cell As Range
Dim cell2 As Range

Dim top_row As Range
Dim match_top_row As Range
Dim match_top_row_first As Range


Dim top_row_num As Long
Dim init_rule_top As Range
Dim init_rule_range As Range
Dim sub_rule_range As Range
Dim rule_range As Range
Dim or_breakout As Range
Dim rule_one As Range
Dim categories As Range
Dim all_or_zeros As VbMsgBoxResult

all_or_zeros = MsgBox("Run for zeros only?", vbYesNo)

start_time = Timer

Set active_workbook = ActiveWorkbook

counter = 0
counter2 = 0

For Each sht In active_workbook.Worksheets
    sht.Activate
    text_to_columns
    If sht.Name = "Data Audit" Or sht.Name = "Certifications" Or sht.Name = "Team Member List" Then
        counter = counter + 1
    End If
Next
        
Workbooks.Open "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Automation\Reference Files - DO NOT CHANGE\Practice File - For Matching.xlsx"
Set opened_workbook = ActiveWorkbook
        
For Each sht In opened_workbook.Worksheets
    If counter2 > 0 Then
        Exit For
    End If
    If sht.Name = "Team Member List" Then
        active_workbook.Activate
        Sheets(sht.Name).Activate
        Set match_top_row_first = Columns(1).Find("Team Member ID", LookAt:=xlWhole)
        Set match_top_row = Range(match_top_row_first, match_top_row_first.End(xlToRight))
        For Each cell In match_top_row
            If cell.Value <> sht.cells(cell.row, cell.Column).Value Then
                counter2 = counter2 + 1
                MsgBox cell.Value & " <> " & sht.cells(cell.row, cell.Column).Value
                Exit For
            End If
        Next
    ElseIf sht.Name = "Data Audit" Then
        active_workbook.Activate
        Sheets(sht.Name).Activate
        Set match_top_row_first = Columns(1).Find("Compensation Grade", LookAt:=xlWhole)
        Set match_top_row = Range(match_top_row_first, match_top_row_first.End(xlToRight))
        For Each cell In match_top_row
            If cell.Value <> sht.cells(cell.row, cell.Column).Value Then
                counter2 = counter2 + 1
                MsgBox cell.Value + " <> " + sht.cells(cell.row, cell.Column).Value
                Exit For
            End If
        Next
    ElseIf sht.Name = "Certifications" Then
        active_workbook.Activate
        Sheets(sht.Name).Activate
        Set match_top_row_first = Columns(1).Find("Preferred Name", LookAt:=xlWhole)
        Set match_top_row = Range(match_top_row_first, match_top_row_first.End(xlToRight))
        For Each cell In match_top_row
            If cell.Value <> sht.cells(cell.row, cell.Column).Value Then
                counter2 = counter2 + 1
                MsgBox cell.Value + " <> " + sht.cells(cell.row, cell.Column).Value
                Exit For
            End If
        Next
    End If
Next


                
                
opened_workbook.Close

active_workbook.Activate


'MsgBox counter & " is the value of counter 1, it should be 3 if they match"
'MsgBox counter2 & " is the value of counter 2, it should be 0 if they match"
    
If counter <> 3 Or counter2 <> 0 Then
    MsgBox "Please provide correct data files.", vbOKOnly, "DANGER WILL ROBINSON!!"
Else
    Sheets("Data Audit").Activate
    top_row_num = Columns(1).Find("Compensation Grade", LookAt:=xlWhole).row

    ClearOut top_row_num

    Sheets("Data Audit").Activate
    Set init_rule_top = Rows(top_row_num).Find("Gr Profile: Eligibility Rules", LookAt:=xlWhole)
    Set init_rule_range = Range(init_rule_top.Offset(1, 0), init_rule_top.End(xlDown))

    FixDataIssues init_rule_range

    Set sub_rule_range = init_rule_range.Offset(0, 1)
    FirstOrLogic top_row_num, sub_rule_range, init_rule_range
    AndLogic top_row_num

    Set rule_range = Range(Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole).Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column))

    Set or_breakout = rule_range.End(xlUp).End(xlToRight).Offset(0, 5)

    Set rule_one = Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole)

    counter = 1
    While counter <> 0
       counter = 0
        For Each cell In rule_range
            If InStr(1, cell.Value, " OR ") Then
                counter = counter + 1
            End If
        Next
        If counter > 0 Then
            SecondOrLogic top_row_num, rule_range, or_breakout
            OrCleanUp top_row_num, or_breakout, rule_range, rule_one
        End If
    Wend

    Set rule_range = Range(Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole).Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column))
            
    FixDataIssues rule_range
    
    Set categories = Range(Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(top_row_num).Find("TM Sup Org Location <> ", LookAt:=xlWhole))
    
    MoveCategories rule_range, categories
    

    
    TeamMemberListAndCerts all_or_zeros
    
    CopyTMData
    
    SplitOutResults
    
    HideSheets
    
    tm_count = Sheets("For Looping").Range("A3", Sheets("For Looping").Range("A3").End(xlDown)).Rows.Count
    minutes_elapsed = Round((Timer - start_time) / 60, 2)

    'MsgBox "The macro finished running in " & minutes_elapsed & " minutes. " & tm_count & " team members were assessed."
    
    'ActiveWorkbook.SaveAs "I:\Human Resources\Strategy and Planning Team\The Vinceinerators\Kirk\Raw Data Dump\Eligibility Rule Audit Post-Run - 1.21.20 V2.xlsx"
    'ActiveWorkbook.Close
    
End If


End Sub




Sub text_to_columns()

Dim rng As Range

Range("A10").Select
Set rng = Selection.CurrentRegion
rng.UnMerge

cells.SpecialCells(xlCellTypeLastCell).Offset(0, 1).Copy

rng.PasteSpecial xlPasteValues, xlPasteSpecialOperationAdd



End Sub


Sub check_no_matches()

'NOT CURRENTLY BEING USED

Dim rng As Range
Dim row_rng As Range


Sheets("No Matches").Activate
Set rng = Range("A3", cells(Range("A3").End(xlDown).row, Range("A2").End(xlToRight).Column))

For Each row_rng In rng.Rows
    row_rng.Copy
    Sheets("Data Audit").Activate
    Range("BA3", cells(Range("AP2").End(xlDown).row, Range("BL2").Column)).PasteSpecial xlPasteValues
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
        Range("A2:CH2").AutoFilter Field:=86, Criteria1:=Range("BB2").Value
    Else
        Range("A2:CH2").AutoFilter Field:=86, Criteria1:=Range("BB2").Value
    End If
    
    
    
    

End Sub

Function TeamMemberListAndCerts(all_or_zeros)

Dim data_range As Range
Dim num_cols As Integer
Dim categories As Range
Dim cell As Range
Dim companies As Range
Dim row As Range
Dim data_rows As Range
Dim counter As Integer
Dim top_row_num As Long
Dim min_col As Long
Dim mid_col As Long
Dim max_col As Long
Dim cert_first_cell As Range
Dim zero_categories As Range
Dim leader_location As Double

Dim last_column_offset As Double

Dim location_list As Range
Dim sub_sup_org As Range
Dim sup_org As Range




Application.Calculation = xlCalculationAutomatic


'This section below will identify people with zero values for ranges and then remove everyone else
'Need to figure out how to set this up.
'WHAT WE NEED
'Current Team Member Details with Pay
'Certifications
'Data Audit for Compensation Grades and Grade Profiles


Sheets("Team Member List").Activate
Sheets("Team Member List").Copy Before:=ActiveSheet

ActiveSheet.Name = "Copied TM List"

top_row_num = Columns(1).Find("Team Member ID", LookAt:=xlWhole).row

If all_or_zeros = vbYes Then
    cells(top_row_num, 1).End(xlToRight).Copy
    cells(top_row_num, 1).End(xlToRight).Offset(0, 1).Value = "Has Zero Ranges?"
    cells(top_row_num, 1).End(xlToRight).PasteSpecial xlPasteFormats

    min_col = Rows(top_row_num).Find("Compensation Range - Minimum", LookAt:=xlWhole).Column
    mid_col = Rows(top_row_num).Find("Compensation Range - Midpoint", LookAt:=xlWhole).Column
    max_col = Rows(top_row_num).Find("Compensation Range - Maximum", LookAt:=xlWhole).Column

    cells(top_row_num, 1).End(xlToRight).Offset(1, 0).Select
    Selection.FormulaR1C1 = "=IF(AND(RC" & min_col & "=0,RC" & mid_col & "=0,RC" & max_col & "=0),""Yes"","""")"
    Selection.Copy

    num_cols = Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight)).Columns.Count

    Range(Selection.Offset(1, 0), cells(top_row_num, 1).End(xlDown).Offset(0, num_cols - 1)).PasteSpecial xlPasteFormulas

    ActiveSheet.Calculate

    Set data_range = Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlDown).Offset(0, num_cols - 1))
    If ActiveSheet.AutoFilterMode = True Then
        data_range.AutoFilterMode = False
    End If

    data_range.AutoFilter Field:=Rows(top_row_num).Find("Has Zero Ranges?", LookAt:=xlWhole).Column, Criteria1:="="


    Range(cells(top_row_num, 1).Offset(1, 0), cells(top_row_num, 1).End(xlToRight).End(xlDown)).Delete Shift:=xlUp
    On Error Resume Next
        ActiveSheet.ShowAllData
    On Error GoTo 0
End If

ActiveSheet.AutoFilterMode = False

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "References"

Range("B2").Value = "Certification"
Range("B3").Value = "Company"
Range("B4").Value = "Job Code"
Range("B5").Value = "Job Focus"
Range("B6").Value = "Loan Allocation"
Range("B7").Value = "Loan Purpose"
Range("B8").Value = "Loan Channel"
Range("B9").Value = "Scheduled Hours"
Range("B10").Value = "TM Location"
Range("B11").Value = "TM Sup Org Location"
Range("B12").Value = "Time Type"


Set categories = Range("B2", Range("B2").End(xlDown))

Range("H2").Value = "Amrock Connecticut, Inc."
Range("H3").Value = "Amrock Title California, Inc."
Range("H4").Value = "Amrock Title Insurance Company"
Range("H5").Value = "Amrock, Inc."
Range("H6").Value = "Bedrock Building Services LLC"
Range("H7").Value = "Bedrock Management Services LLC"
Range("H8").Value = "One Reverse Mortgage, LLC"
Range("H9").Value = "Quicken Loans Inc."
Range("H10").Value = "Rock Connections"
Range("H11").Value = "In House Realty Canada ULC"
Range("H12").Value = "Rock FOC Technologies Canada ULC"
Range("H13").Value = "Lendesk"


Range("I2:I5").Value = "Amrock"
Range("I6:I7").Value = "Bedrock"
Range("I8").Value = "One Reverse Mortgage"
Range("I9").Value = "Quicken Loans"
Range("I10").Value = "Rock Connections"
Range("I11:I13").Value = "Canada"


Range("B19").Value = "Detroit, MI"
Range("B20").Value = "Phoenix, AZ"
Range("B21").Value = "Tempe, AZ"
Range("B22").Value = "San Diego, CA"
Range("B23").Value = "Cleveland, OH"
Range("B24").Value = "Cerritos, CA"
Range("B25").Value = "Coraopolis, PA"
Range("B26").Value = "Charlotte, NC"
Range("B27").Value = "Remote - New York"
Range("B28").Value = "Windsor, ON"
Range("B29").Value = "Washington, DC"
Range("B30").Value = "Douglasville, GA"
Range("B31").Value = "Oakland, California"
Range("B32").Value = "Fort Worth, TX"
Range("B33").Value = "Remote - Illinois"
Range("B34").Value = "Moonachie, NJ"
Range("B35").Value = "New Jersey - Moonachie"
Range("B36").Value = "Toronto, ON"
Range("B37").Value = "Los Angeles, CA"
Range("B38").Value = "Remote - Michigan"
Range("B39").Value = "Remote - Washington"
Range("B40").Value = "Dallas, TX"
Range("B41").Value = "Vancouver, BC"
Range("B42").Value = "Mountain View, CA"

Range("C19").Value = "Michigan"
Range("C20").Value = "Arizona"
Range("C21").Value = "Arizona"
Range("C22").Value = "California"
Range("C23").Value = "Ohio"
Range("C24").Value = "California"
Range("C25").Value = "Pennsylvania"
Range("C26").Value = "North Carolina"
Range("C27").Value = "New York"
Range("C28").Value = "Ontario"
Range("C29").Value = "Washington, DC"
Range("C30").Value = "Georgia"
Range("C31").Value = "OakCal"
Range("C32").Value = "Texas"
Range("C33").Value = "Illinois"
Range("C34").Value = "New Jersey"
Range("C35").Value = "New Jersey"
Range("C36").Value = "Ontario"
Range("C37").Value = "California"
Range("C38").Value = "Michigan"
Range("C39").Value = "Washington"
Range("C40").Value = "Texas"
Range("C41").Value = "British Columbia"
Range("C42").Value = "California"


Range("K4").Value = "OakCal"

Range("L4").Value = "California"

If IsEmpty(Range("K5")) Then
    Set sub_sup_org = Range("K4")
Else
    Set sub_sup_org = Range("K4", Range("K4").End(xlDown))
End If

Set sup_org = sub_sup_org.Offset(0, 1)
ActiveWorkbook.Names.Add Name:="SubSupOrg", RefersTo:=sub_sup_org
ActiveWorkbook.Names.Add Name:="SupOrg", RefersTo:=sup_org


Set companies = Range("H2", Range("H2").End(xlToRight).End(xlDown))


ActiveWorkbook.Names.Add Name:="Companies", RefersTo:=companies

For Each cell In categories
    cell.Copy
    Sheets("Copied TM List").Activate
    cells(top_row_num, 1).End(xlToRight).Offset(0, 1).PasteSpecial xlPasteValues
Next

Set zero_categories = Range(Rows(top_row_num).Find("Certification", LookAt:=xlWhole), cells(top_row_num, 1).End(xlToRight))

Sheets("Copied TM List").Activate
cells(top_row_num, 1).Copy
Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight)).PasteSpecial xlPasteFormats

Sheets("Certifications").Activate

Set cert_first_cell = Columns(1).Find("Preferred Name", LookAt:=xlWhole)

cert_first_cell.End(xlToRight).Offset(0, 1).Value = "Active?"
cert_first_cell.End(xlToRight).Offset(1, 0).Value = "=if(isblank(G3),""Yes"",if(G3<TODAY(),""No"",""Yes""))"
cert_first_cell.End(xlToRight).Offset(1, 0).Copy
Range(cert_first_cell.End(xlToRight).Offset(2, 0), cells(cert_first_cell.End(xlDown).row, cert_first_cell.End(xlToRight).Column)).PasteSpecial xlPasteFormulas

ActiveSheet.Calculate

If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If
    
Range(cert_first_cell, cert_first_cell.End(xlToRight).End(xlDown)).AutoFilter Field:=cert_first_cell.End(xlToRight).Column, Criteria1:="No"

Range(cert_first_cell.Offset(1, 0), cert_first_cell.End(xlToRight).End(xlDown)).Delete Shift:=xlUp
ActiveSheet.ShowAllData

ActiveSheet.AutoFilterMode = False

Sheets("Team Member List").Activate
leader_location = Rows(top_row_num).Find("Location Address - State/Province", LookAt:=xlWhole).Column

'THIS NEEDS TO BE FACTORED IN FOR LONG TERM USE
'GOING TO BUILD OUT REFERENCES WORKBOOK AND JUST MOVE THE SHEET OVER
'What i mean, is that i will create a workbook that is saved to a shared folder. That will have the references in it.


Sheets("References").Activate



Set location_list = Range("B19", Range("B19").End(xlDown))

Sheets("Copied TM List").Activate


For Each cell In Range(Rows(top_row_num).Find("Supervisory Organization - Primary Location", LookAt:=xlWhole).Offset(1, 0), Rows(top_row_num).Find("Supervisory Organization - Primary Location", LookAt:=xlWhole).End(xlDown))
    cell.Value = location_list.Find(cell.Value).Offset(0, 1)
Next


Sheets("Copied TM List").Activate

If all_or_zeros = vbYes Then
    Rows(top_row_num).Find("Has Zero Ranges?", LookAt:=xlWhole).EntireColumn.Delete Shift:=xlToLeft
End If

'Certification
zero_categories.Find("Certification", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC" & Rows(top_row_num).Find("Name (First Last)", LookAt:=xlWhole).Column & ",Certifications!C1:C1,1,0),"""")="""","""",""Yes"")"
'Company
zero_categories.Find("Company", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IFERROR(VLOOKUP(RC" & Rows(top_row_num).Find("Company - Name", LookAt:=xlWhole).Column & ",Companies,2,0),RC" & Rows(top_row_num).Find("Company - Name", LookAt:=xlWhole).Column & ")"
'Job Code
zero_categories.Find("Job Code", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Job Code", LookAt:=xlWhole).Column
'Job Focus
zero_categories.Find("Job Focus", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Job Focus", LookAt:=xlWhole).Column
'Loan Allocation
zero_categories.Find("Loan Allocation", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Specialty Allocation", LookAt:=xlWhole).Column
'Loan Purpose
zero_categories.Find("Loan Purpose", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Specialty Purpose", LookAt:=xlWhole).Column
'Loan Channel
zero_categories.Find("Loan Channel", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Specialty Channel", LookAt:=xlWhole).Column
'Sheduled Hours
zero_categories.Find("Scheduled Hours", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Scheduled Weekly Hours", LookAt:=xlWhole).Column
'TM Location - THIS IS WHAT I WOULD NEED TO UPDATE IF MITCH CHANGES IT
zero_categories.Find("TM Location", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(top_row_num).Find("Location Address - State/Province", LookAt:=xlWhole).Column

'TM Sup Org Loaction
zero_categories.Find("TM Sup Org Location", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = _
        "=RC" & Rows(top_row_num).Find("Supervisory Organization - Primary Location", LookAt:=xlWhole).Column
'Time Type
zero_categories.Find("Time Type", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = _
        "=RC" & Rows(top_row_num).Find("Time Type", LookAt:=xlWhole).Column
    



last_column_offset = zero_categories.Find("Time Type", LookAt:=xlWhole).Column - Rows(top_row_num).Find("Supervisory Organization - Primary Location", LookAt:=xlWhole).Column

Range(Rows(top_row_num).Find("Certification", LookAt:=xlWhole).Offset(1, 0), Rows(top_row_num).Find("Certification", LookAt:=xlWhole).End(xlToRight).Offset(1, 0)).Copy
Range(Rows(top_row_num).Find("Certification", LookAt:=xlWhole).Offset(2, 0), Rows(top_row_num).Find("Supervisory Organization - Primary Location", LookAt:=xlWhole).End(xlDown).Offset(0, last_column_offset)).PasteSpecial xlPasteFormulas

'^^ need to add calculation for the offset of columns to be the last column minus the "employment status" column to get the correct offset


ActiveSheet.Calculate

Sheets("Copied TM List").Copy Before:=ActiveSheet
ActiveSheet.Name = "For Looping"

Sheets("For Looping").Activate


cells.Select
With Selection
    .Copy
    .PasteSpecial xlPasteValues
End With

Range("A1", cells(top_row_num, 1).Offset(-2, 0)).EntireRow.Delete Shift:=xlUp


Range(Rows(2).Find("Compensation Grade Profile", LookAt:=xlWhole), Rows(2).Find("Certification", LookAt:=xlWhole).Offset(0, -1)).EntireColumn.Delete Shift:=xlToLeft
'^ Needs to be column "Compensation Grade Profile" to "Certification" column - 1

Range(Rows(2).Find("Name (Last, First)", LookAt:=xlWhole), Rows(2).Find("DRIVE Eligibility", LookAt:=xlWhole)).EntireColumn.Delete Shift:=xlToLeft
' Name (Last, First) to DRIVE Eligibility

Sheets("For Looping").Activate

Set data_rows = Range("A3", cells(Range("A2").End(xlDown).row, Range("A2").End(xlToRight).Column))

For Each cell In data_rows
    If InStr(cell.Value, "Not Applicable") Or InStr(cell.Value, "No Specialization") Then
        cell.ClearContents
    End If
    If Len(cell) = 0 Then
        cell.ClearContents
    End If
Next

If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If

Range("A2", Range("A2").End(xlToRight)).AutoFilter Field:=2, Criteria1:="="
Range(Range("A2").Offset(1, 0), Range("A2").End(xlDown)).SpecialCells(xlCellTypeVisible).EntireRow.Delete Shift:=xlUp

ActiveSheet.AutoFilterMode = False






    

End Function

Function CopyTMData()

Dim counter As Integer
Dim row As Range
Dim data_rows As Range
Dim title_row As Range
Dim cell As Range
Dim calc1 As String
Dim calc2 As String
Dim counter2 As Integer
Dim tm_data_start As Range
Dim dist_to_compare As Double
Dim dist_to_tm As Double
Dim tm_data_top As Range
Dim dist_to_tm2 As Double
Dim tm_paste_range As Range
Dim filter_range As Range

Dim start_time As Double
Dim time_counter As Double
Dim perc_notify As Double
Dim perc_complete As Double
Dim perc_counter As Double




Dim data_top_row_num As Double

Sheets("For Looping").Activate

Set title_row = Range(Columns(1).Find("Team Member ID", LookAt:=xlWhole), Columns(1).Find("Team Member ID", LookAt:=xlWhole).End(xlToRight))

Sheets("Data Audit").Activate
data_top_row_num = Columns(1).Find("Compensation Grade", LookAt:=xlWhole).row

Set tm_data_start = cells(data_top_row_num, 1).End(xlToRight).Offset(0, 5)

For Each cell In title_row
    cell.Copy
    Sheets("Data Audit").Activate
    If IsEmpty(tm_data_start) Then
        tm_data_start.PasteSpecial xlPasteValues
    ElseIf IsEmpty(tm_data_start.Offset(0, 1)) Then
        tm_data_start.Offset(0, 1).PasteSpecial xlPasteValues
    Else
        tm_data_start.End(xlToRight).Offset(0, 1).PasteSpecial xlPasteValues
    End If
Next

Application.Calculation = xlCalculationAutomatic

Sheets("Data Audit").Activate
Range(Rows(data_top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(data_top_row_num).Find("Time Type <> ", LookAt:=xlWhole)).Copy
tm_data_start.End(xlToRight).Offset(0, 1).PasteSpecial xlPasteValues
tm_data_start.End(xlToRight).Offset(0, 1).Value = "Match?"
tm_data_start.End(xlToRight).Offset(0, 1).Value = "Grade Profile"

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Matches"
Sheets("Data Audit").Activate
Range(tm_data_start, tm_data_start.End(xlToRight)).Copy
Sheets("Matches").Activate
Range("A2").PasteSpecial xlPasteValues

Sheets("Data Audit").Activate

dist_to_compare = Rows(data_top_row_num).Find("Certification = ", LookAt:=xlWhole).Column - Range(tm_data_start, tm_data_start.End(xlToRight)).Find("Certification = ", LookAt:=xlWhole).Column
dist_to_tm = Range(tm_data_start, tm_data_start.End(xlToRight)).Find("Certification", LookAt:=xlWhole).Column - Range(tm_data_start, tm_data_start.End(xlToRight)).Find("Certification = ", LookAt:=xlWhole).Column


Set tm_data_top = Range(tm_data_start, tm_data_start.End(xlToRight))

dist_to_tm2 = tm_data_top.Find("Certification", LookAt:=xlWhole).Column - tm_data_top.Find("Certification <> ", LookAt:=xlWhole).Column

Range(tm_data_start, tm_data_start.End(xlToRight)).Find("Certification = ", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(ISBLANK(RC[" & dist_to_compare & "]),"""",IF(RC[" & dist_to_compare & "]=RC[" & dist_to_tm & "],""Yes"",""No""))"
Range(tm_data_start, tm_data_start.End(xlToRight)).Find("Certification = ", LookAt:=xlWhole).Offset(1, 0).Copy

Range(tm_data_top.Find("Company = ", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Loan Channel = ", LookAt:=xlWhole).Offset(1, 0)).PasteSpecial xlPasteFormulas
Range(tm_data_top.Find("TM Location = ", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Time Type = ", LookAt:=xlWhole).Offset(1, 0)).PasteSpecial xlPasteFormulas


'Range("BT3").Value = "=IF(ISBLANK(AC3),"""",IF(BJ3>=AC3,""Yes"",""No""))"

'Current ONE
tm_data_top.Find("Scheduled Hours >= ", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = _
        "=if(isblank(RC[" & dist_to_compare & _
        "]), """", if(RC[" & dist_to_tm & "]>=RC[" & _
        dist_to_compare & "],""Yes"",""No""))"


'For Certification <>
tm_data_top.Find("Certification <> ", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(ISBLANK(RC[" & dist_to_compare & "]),"""",if(ISBLANK(RC[" & dist_to_tm2 & _
        "]),""Yes"",IF(OR(ISNUMBER(SEARCH(RC[" & dist_to_tm2 & "],RC[" & dist_to_compare & "])),ISNUMBER(SEARCH(XLOOKUP(RC[" & dist_to_tm2 & "],SubSupOrg,SupOrg),RC[" & dist_to_compare & "]))),""No"",""Yes"")))"

tm_data_top.Find("Certification <> ", LookAt:=xlWhole).Offset(1, 0).Copy

Range(tm_data_top.Find("Company <> ", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Loan Channel <> ", LookAt:=xlWhole).Offset(1, 0)).PasteSpecial xlPasteFormulas
Range(tm_data_top.Find("TM Location <> ", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Time Type <> ", LookAt:=xlWhole).Offset(1, 0)).PasteSpecial xlPasteFormulas


tm_data_top.Find("Scheduled Hours < ", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(ISBLANK(RC[" & dist_to_compare & "]), """", IF(RC[" & dist_to_tm2 & "]<RC[" & dist_to_compare & "], ""Yes"",""No""))"

tm_data_top.Find("Match?", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(AND(RC1=RC" & tm_data_top.Find("Compensation Grade", LookAt:=xlWhole).Column & ",COUNTIFS(RC" & tm_data_top.Find("Certification = ", LookAt:=xlWhole).Column & _
        ":RC" & tm_data_top.Find("Time Type <> ", LookAt:=xlWhole).Column & ",""Yes"")>0, COUNTIFS(RC" & tm_data_top.Find("Certification = ", LookAt:=xlWhole).Column & _
        ":RC" & tm_data_top.Find("Time Type <> ", LookAt:=xlWhole).Column & ",""No"")=0),""Yes"","""")"

tm_data_top.Find("Grade Profile", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=RC" & Rows(data_top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole).Column
Range(tm_data_top.Find("Certification = ", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Grade Profile", LookAt:=xlWhole).Offset(1, 0)).Copy





Range(tm_data_top.Find("Certification = ", LookAt:=xlWhole).Offset(2, 0), cells(Rows(data_top_row_num).Find("Rule 1", LookAt:=xlWhole).End(xlDown).row, tm_data_top.Find("Grade Profile", LookAt:=xlWhole).Column)).PasteSpecial xlPasteFormulas


If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If

Set filter_range = Range(cells(data_top_row_num, 1), cells(data_top_row_num, tm_data_top.Find("Grade Profile", LookAt:=xlWhole).Column))
filter_range.AutoFilter
    
Sheets("For Looping").Activate

Set data_rows = Range("A3", cells(Range("A3").End(xlDown).row, Range("A2").End(xlToRight).Column))

counter = 1

Sheets("Data Audit").Activate

Set tm_paste_range = Range(tm_data_top.Find("Team Member ID", LookAt:=xlWhole).Offset(1, 0), cells(Rows(data_top_row_num).Find("Rule 1", LookAt:=xlWhole).End(xlDown).row, tm_data_top.Find("Time Type", LookAt:=xlWhole).Column))

start_time = Timer
perc_notify = 25
time_counter = 0
perc_counter = 0

For Each row In data_rows.Rows
    row.Copy
    Sheets("Data Audit").Activate
    tm_paste_range.PasteSpecial xlPasteValues
    ActiveSheet.Calculate
    counter2 = 0
    For Each cell In Range(tm_data_top.Find("Match?", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Match?", LookAt:=xlWhole).End(xlDown))
        If cell.Value = "Yes" Then
            counter2 = counter2 + 1
        End If
    Next
    If counter2 > 0 Then
        filter_range.AutoFilter Field:=Rows(data_top_row_num).Find("Match?", LookAt:=xlWhole).Column, Criteria1:="Yes"
        If IsEmpty(cells(data_top_row_num, 1).Offset(1, 0)) = False Then
            Range(tm_data_top.Find("Team Member ID", LookAt:=xlWhole).Offset(1, 0), tm_data_top.Find("Grade Profile", LookAt:=xlWhole).End(xlDown)).Copy
            Sheets("Matches").Activate
            If IsEmpty(Range("A3")) Then
                Range("A3").PasteSpecial xlPasteValues
            Else
                Range("A2").End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
            End If
        End If
        Sheets("Data Audit").Activate
        ActiveSheet.ShowAllData
    End If
    tm_paste_range.ClearContents
    time_counter = time_counter + 1
    If time_counter = Round(data_rows.Rows.Count * (perc_notify / 100), 0) Then
        perc_counter = perc_counter + 1
        minutes_elapsed = Round((Timer - start_time) / 60, 2)
        'send_email "kwaters004@gmail.com; kirkwaters@quickenloans.com", "", minutes_elapsed & " minutes elapsed, " & (perc_counter * perc_notify) & "% Complete: ~" & (100 / (perc_counter * perc_notify)) * minutes_elapsed - minutes_elapsed & " mintues remaining"
        time_counter = 0
    End If
Next
Sheets("Matches").Activate

If WorksheetFunction.CountA(Range("A3", Range("A3").End(xlDown))) <> 0 Then
    Range("A2").End(xlToRight).Offset(0, 1).Value = "Remove Dupes"
    Range("A2").End(xlToRight).Offset(1, 0).FormulaR1C1 = "=COUNTIFS(R3C1:RC1,RC1,R3C" & Rows(2).Find("Grade Profile", LookAt:=xlWhole).Column & _
        ":RC" & Rows(2).Find("Grade Profile", LookAt:=xlWhole).Column & ", RC" & Rows(2).Find("Grade Profile", LookAt:=xlWhole).Column & ")"
    If IsEmpty(Range("A2").End(xlToRight).Offset(2, 0)) = True Then
        Range("A2").End(xlToRight).Offset(1, 0).Copy
        Range(Range("A2").End(xlToRight).Offset(2, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).PasteSpecial xlPasteFormulas
        ActiveSheet.Calculate
        Range(Range("A2").End(xlToRight).Offset(1, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).Copy
        Range(Range("A2").End(xlToRight).Offset(1, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).PasteSpecial xlPasteValues
        While Application.WorksheetFunction.Sum(Range(Range("A2").End(xlToRight).Offset(1, 0), _
                Range("A2").End(xlToRight).End(xlDown))) <> Range(Range("A2").End(xlToRight).Offset(1, 0), Range("A2").End(xlToRight).End(xlDown)).Rows.Count
            For Each cell In Range(Range("A2").End(xlToRight).Offset(1, 0), Range("A2").End(xlToRight).End(xlDown))
                If cell.Value > 1 Then
                    cell.EntireRow.Delete Shift:=xlUp
                End If
            Next
        Wend
    End If
    'Range("AI1").EntireColumn.Delete Shift:=xlToLeft
End If



End Function

Sub SplitOutResults()

Sheets("Matches").Activate
Range("A2").End(xlToRight).Offset(0, 1).Value = "Number of Matches"
Range("A2").End(xlToRight).Offset(1, 0).Value = "=COUNTIFS($A:$A,A3)"
Range("A2").End(xlToRight).Offset(1, 0).Copy
Range(Range("A2").End(xlToRight).Offset(2, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).PasteSpecial xlPasteFormulas
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If
Range("A2", Range("A2").End(xlToRight)).AutoFilter Field:=Rows(2).Find("Number of Matches", LookAt:=xlWhole).Column, Criteria1:="1"

Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Single Matches"

Sheets("Single Matches").Activate
Range("A2").PasteSpecial xlPasteValues

Sheets("Matches").Activate
ActiveSheet.ShowAllData
Range("A2", Range("A2").End(xlToRight)).AutoFilter Field:=Rows(2).Find("Number of Matches", LookAt:=xlWhole).Column, Criteria1:=">1"
Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "Multiple Matches"

Sheets("Multiple Matches").Activate
Range("A2").PasteSpecial xlPasteValues

Sheets("For Looping").Activate
Range("A2").End(xlToRight).Offset(1, 1).Value = "=IFERROR(VLOOKUP(A3,Matches!$A:$A,1,0),"""")"
Range("A2").End(xlToRight).Offset(1, 1).Copy
Range(Range("A2").End(xlToRight).Offset(2, 1), cells(Range("A2").End(xlDown).row, Range("A2").End(xlToRight).Offset(0, 1).Column)).PasteSpecial xlPasteFormulas
Range("A2", Range("A2").End(xlToRight).Offset(0, 1)).AutoFilter Field:=Range("A2").End(xlToRight).Offset(0, 1).Column, Criteria1:="="
Range(Range("A2"), cells(Range("A2").End(xlDown).row, Range("A2").End(xlToRight).Offset(0, 1).Column)).Copy

Sheets.Add After:=ActiveSheet
ActiveSheet.Name = "No Matches"

Sheets("No Matches").Activate
Range("A2").PasteSpecial xlPasteValues

Sheets("For Looping").Activate
ActiveSheet.ShowAllData



End Sub

Sub HideSheets()

Dim sht As Worksheet

For Each sht In ActiveWorkbook.Worksheets
    If sht.Name <> "No Matches" And sht.Name <> "Multiple Matches" And sht.Name <> "Single Matches" Then
        sht.Visible = xlSheetHidden
    End If
Next




End Sub



Function ClearOut(top_row_num)

Dim data_audit As Range, data_audit_top_row As Range
Dim grade_prof_col As Range
Dim grade_prof_top As Range
'Dim top_row_num As Long
Dim new_profiles As Range
Dim cell As Range
Dim n As Range




Application.Calculation = xlCalculationAutomatic

cells.Select
Selection.UnMerge

'top_row_num = Columns(1).Find("Compensation Grade", Lookat:=xlWhole).row
Set grade_prof_top = Rows(top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole)
Set grade_prof_col = Range(grade_prof_top.Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, grade_prof_top.Column))

'For Each cell In grade_prof_col
'    If InStr(cell.Value, "inactive") Then
'     cell.Value = ""
'    End If
'Next

Set data_audit_top_row = Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight))
Set data_audit = Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlDown).Offset(0, data_audit_top_row.Columns.Count))

  
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If
    data_audit.AutoFilter Field:=cells(top_row_num, Rows(top_row_num).Find("Grade Profile Inactive", LookAt:=xlWhole).Column).Column, Criteria1:="=", Criteria2:="Yes", Operator:=xlOr

Range(cells(top_row_num, 1).Offset(1, 0), cells(top_row_num, 1).End(xlDown)).EntireRow.Delete Shift:=xlUp

ActiveSheet.ShowAllData

On Error Resume Next
    Sheets("New Grade Profiles").Activate
If Err.Number = 0 Then
    Range("D3").FormulaR1C1 = "=RC1 & "": "" & RC2"
    Range("D3").Copy
    Range("D3", Range("C2").End(xlDown).Offset(0, 1)).PasteSpecial xlPasteFormulas
    Range("D3", Range("C2").End(xlDown).Offset(0, 1)).Copy
    Range("D3").PasteSpecial xlPasteValues
    Set new_profiles = Range("D3", Range("C2").End(xlDown).Offset(0, 1))
    Sheets("Data Audit").Activate
    For Each n In new_profiles
        For Each cell In Range(Rows(top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole).Offset(1, 0), _
                Rows(top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole).End(xlDown))
            If n.Value = cell.Value Then
                cell.EntireRow.Delete Shift:=xlUp
                Exit For
            End If
        Next
    Next
    Sheets("New Grade Profiles").Activate
    Sheets("New Grade Profiles").Range("A3", Range("A2").End(xlDown)).Copy
    Sheets("Data Audit").cells(top_row_num, 1).End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
    Sheets("New Grade Profiles").Range("D3", Range("C2").End(xlDown).Offset(0, 1)).Copy
    Sheets("Data Audit").cells(top_row_num, grade_prof_top.Column).End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
    Sheets("New Grade Profiles").Range("C3", Range("C2").End(xlDown)).Copy
    Sheets("Data Audit").Rows(top_row_num).Find("Gr Profile: Eligibility Rules", LookAt:=xlWhole).End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
End If
Err.Clear
On Error GoTo 0

End Function

Function FixDataIssues(rule_range)


Dim cell As Range


Application.Calculation = xlCalculationAutomatic

Sheets("Data Audit").Activate


For Each cell In rule_range
    Trim (cell.Value)
    cell = replace(cell.Value, "Team Member Location", "TM Location")
    cell = replace(cell.Value, "(DE, LAPP, SARS)", "Yes")
    cell = replace(cell.Value, "DE, LAPP, SARS", "Yes")
    cell = replace(cell.Value, "not in", "<>")
    cell = replace(cell.Value, "is in", "=")
    cell = replace(cell.Value, "(AZ, NC)", "Arizona OR North Carolina")
    cell = replace(cell.Value, "is empty", "<> Yes")
    cell = replace(cell.Value, " and ", " AND ")
    cell = replace(cell.Value, " or ", " OR ")
    cell = replace(cell.Value, "Full Time- ", "")
    cell = replace(cell.Value, "Part Time- ", "")
    If InStr(cell.Value, "TM Location") = 0 And InStr(cell.Value, "Allocation") = 0 And InStr(cell.Value, "TM Sup Org Location") = 0 Then
        cell = replace(cell.Value, "Location", "TM Location")
    End If
    cell = replace(cell.Value, "Loan Specialty Allocation", "Loan Allocation")
    cell = replace(cell.Value, "Loan Specialty Purpose", "Loan Purpose")
    cell = replace(cell.Value, "Loan Specialty Channel", "Loan Channel")
    cell = replace(cell.Value, "One Reverse Mortgage LLC", "One Reverse Mortgage")
    cell = replace(cell.Value, "One Reverse Mortgage LLC", "One Reverse Mortgage")
    cell = replace(cell.Value, "CA", "California")
    cell = replace(cell.Value, "MI", "Michigan")
    cell = replace(cell.Value, "AZ", "Arizona")
    cell = replace(cell.Value, "NC", "North Carolina")
    cell = replace(cell.Value, "Scheduled Hours = ", "Scheduled Hours >= ")
    cell = replace(cell.Value, "Oakland, California", "OakCal")
    cell = replace(cell.Value, "Washington, DC", "WashDC")
    If InStr(cell.Value, "Refinance") = 0 Then
        cell = replace(cell.Value, "Refi", "Refinance")
    End If
    
Next
    
End Function




Function FirstOrLogic(top_row_num, sub_rule_range, init_rule_range)

Dim num As Integer
Dim num_new_sheets As Integer
'Dim top_row_num As Long
Dim sub_row_top As Long
Dim sub_row_bottom As Long
Dim sub_col As Long




Application.Calculation = xlCalculationAutomatic

Sheets("Data Audit").Activate

sub_rule_range.FormulaR1C1 = "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(RC[-1],"") OR ("",""/""),"")"",""""),""("","""")"

sub_rule_range.Copy
sub_rule_range.PasteSpecial xlPasteValues

sub_rule_range.TextToColumns Destination:=sub_rule_range.item(1), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

cells(1, sub_rule_range.Column).FormulaR1C1 = "=COUNTIFS(R[" & top_row_num + 1 & "]C:R1000C,""<>""&"""")"
cells(1, sub_rule_range.Column).Copy
Range(cells(1, sub_rule_range.Column), cells(1, sub_rule_range.Column).Offset(0, 5)).PasteSpecial xlPasteFormulas

If cells(1, sub_rule_range.Column).Offset(0, 1).Value <> 0 Then
    For Each cell In Range(cells(1, sub_rule_range.Column), cells(1, sub_rule_range.Column).End(xlToRight))
        If cell.Value = 0 Then
            cell.ClearContents
        End If
    Next
End If
num = 1

For Each cell In Range(cells(2, sub_rule_range.Column), cells(1, sub_rule_range.Column).End(xlToRight).Offset(1, 0))
    cell.Value = num
    num = num + 1
Next

num_new_sheets = Range(cells(1, sub_rule_range.Column), cells(1, sub_rule_range.Column).End(xlToRight)).Columns.Count
num = 1

While num <> num_new_sheets
    Sheets("Data Audit").Activate
    Range(cells(2, sub_rule_range.Column), cells(2, sub_rule_range.Column).End(xlDown).Offset(0, num - 1)).EntireColumn.Hidden = True
    Range(cells(top_row_num, 1), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, sub_rule_range.Column).Offset(0, num).Column)).SpecialCells(xlCellTypeVisible).Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Or Sheet " & num
    Range("A2").PasteSpecial xlPasteValues
    num = num + 1
Wend
   
Sheets("Data Audit").Activate

cells.EntireColumn.Hidden = False

Range(cells(top_row_num, sub_rule_range.Offset(0, 1).Column), cells(top_row_num, sub_rule_range.Offset(0, 1).Column).End(xlToRight)).EntireColumn.Delete Shift:=xlToLeft

num = 1

'These fixed ranges are ok because we are creating the sheets that we are referencing.

While num <> num_new_sheets
    Sheets("Or Sheet " & num).Activate
    If IsEmpty(Range("A4")) Then
        Range("A3", Range("A2").End(xlToRight).Offset(1, 0)).Copy
    Else
        Range("A3", cells(Range("A3").End(xlDown).row, Range("A2").End(xlToRight).Column)).Copy
    End If
    Sheets("Data Audit").Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
    num = num + 1
Wend

Sheets("Data Audit").Activate

sub_col = sub_rule_range.Column


If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
    Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight)).AutoFilter Field:=sub_col, Criteria1:="="
    Range(cells(top_row_num, 1).Offset(1, 0), cells(top_row_num, 1).End(xlDown)).EntireRow.Delete Shift:=xlUp
Else
    Range(cells(top_row_num, 1), cells(top_row_num, 1).End(xlToRight)).AutoFilter Field:=sub_col, Criteria1:="="
    Range(cells(top_row_num, 1).Offset(1, 0), cells(top_row_num, 1).End(xlDown)).EntireRow.Delete Shift:=xlUp
End If

ActiveSheet.ShowAllData

Application.DisplayAlerts = False

Range(cells(top_row_num + 1, sub_rule_range.Column), cells(top_row_num, sub_rule_range.Column).End(xlDown)).Copy



init_rule_range.item(1).PasteSpecial xlPasteValues

sub_rule_range.EntireColumn.ClearContents



num = 1

While num <> num_new_sheets
    Sheets("Or Sheet " & num).Delete
    num = num + 1
Wend

Application.DisplayAlerts = True

End Function
Function AndLogic(top_row_num)

Dim cell As Range
Dim rule_top_row As Range
Dim rule_range As Range
Dim counter As Integer
Dim category_list As Variant
Dim n As Long
Dim sub_and As Range
Dim sub_and_offset As Integer




Application.Calculation = xlCalculationAutomatic

Sheets("Data Audit").Activate

category_list = Array("Certification = ", "Company = ", "Job Code = ", "Job Focus = ", "Loan Allocation = ", _
        "Loan Purpose = ", "Loan Channel = ", "Scheduled Hours >= ", "TM Location = ", "TM Sup Org Location = ", "Time Type = ", _
        "Certification <> ", "Company <> ", "Job Code <> ", "Job Focus <> ", "Loan Allocation <> ", _
        "Loan Purpose <> ", "Loan Channel <> ", "Scheduled Hours < ", "TM Location <> ", "TM Sup Org Location <> ", "Time Type <> ", "Sub AND")

For n = 0 To UBound(category_list)
    cells(top_row_num, 1).End(xlToRight).Offset(0, 1).Value = category_list(n)
Next n

sub_and_offset = Rows(top_row_num).Find("Gr Profile: Eligibility Rules", LookAt:=xlWhole).Column - cells(top_row_num, 1).End(xlToRight).Column

Range(cells(top_row_num, 1).End(xlToRight).Offset(1, 0), _
        cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column)).FormulaR1C1 = _
        "=SUBSTITUTE(SUBSTITUTE(RC[" & sub_and_offset & "],"" AND "",""/""),""("","""")"

'START OF AND PARSING. WILL NEED "OR" PARSING FOR MULTIPLE RULE ELIGIBILITY RULES


Range(cells(top_row_num, 1).End(xlToRight).Offset(1, 0), cells(top_row_num, 1).End(xlToRight).End(xlDown)).Copy
Application.DisplayAlerts = False
cells(top_row_num, 1).End(xlToRight).Offset(1, 0).PasteSpecial xlPasteValues
Application.DisplayAlerts = True
   
Range(cells(top_row_num, 1).End(xlToRight).Offset(1, 0), _
        cells(top_row_num, 1).End(xlToRight).End(xlDown)).TextToColumns Destination:=cells(top_row_num, 1).End(xlToRight).Offset(1, 0), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

'Right here

Set sub_and = Rows(top_row_num).Find("Sub AND", LookAt:=xlWhole)

Range(cells(1, sub_and.Column), cells(1, sub_and.Column).Offset(0, 20)).FormulaR1C1 = "=COUNTIFS(R" & top_row_num + 1 & "C:R10000C,""<>""&"""")"

ActiveSheet.Calculate

For Each cell In Range(cells(1, sub_and.Column), cells(1, sub_and.Column).Offset(0, 20))
    If cell.Value = 0 Then
        cell.ClearContents
    End If
Next

counter = 1
Set rule_top_row = Range(sub_and, cells(1, sub_and.Column).End(xlToRight).Offset(top_row_num - 1, 0))

For Each cell In rule_top_row
    cell.Value = "Rule " & counter
    counter = counter + 1
Next





End Function

Function SecondOrLogic(top_row_num, rule_range, or_breakout)

Dim cell As Range
'Dim rule_range As Range
Dim categories As Range
Dim rule_row_count As Integer
Dim rule_row As Range
Dim no_cat_counter As Integer
Dim cat2 As Range
Dim cat As Range
'Dim or_breakout As Range


Application.Calculation = xlCalculationAutomatic


'Set rule_row = Range("AQ2", Range("AQ2").End(xlToRight))
'rule_row_count = rule_row.Columns.Count

'Set rule_range = Range("AQ3", Range("AQ3").End(xlDown).Offset(0, rule_row_count - 1))
Set categories = Range(Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(top_row_num).Find("Time Type <> ", LookAt:=xlWhole))

For Each cell In rule_range
    If InStr(cell.Value, " OR ") Then
        For Each cat In categories
            If InStr(Left(cell.Value, InStr(cell.Value, " OR ")), cat) Then
                If IsEmpty(cells(cell.row, Range("BA1").Column)) Then
                    cells(cell.row, or_breakout.Column).Value = Left(cell.Value, InStr(cell.Value, " OR ") - 1)
                ElseIf IsEmpty(cells(cell.row, or_breakout.Column).Offset(0, 1)) Then
                    cells(cell.row, or_breakout.Column).Offset(0, 1).Value = Left(cell.Value, InStr(cell.Value, " OR ") - 1)
                Else
                    cells(cell.row, or_breakout.Column).End(xlToRight).Offset(0, 1).Value = Left(cell.Value, InStr(cell.Value, " OR ") - 1)
                End If
                               
            'This IF is checking if the same category shows on both sides of the OR
                If InStr(Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ")), cat) Then
                    If IsEmpty(cells(cell.row, or_breakout.Column)) Then
                        cells(cell.row, or_breakout.Column).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                    ElseIf IsEmpty(cells(cell.row, or_breakout.Column).Offset(0, 1)) Then
                        cells(cell.row, or_breakout.Column).Offset(0, 1).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                    Else
                        cells(cell.row, or_breakout.Column).End(xlToRight).Offset(0, 1).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                    End If
                Else
                    no_cat_counter = 0
                    For Each cat2 In categories
                        
                        If InStr(Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ")), cat2) Then
                            If IsEmpty(cells(cell.row, or_breakout.Column)) Then
                                cells(cell.row, or_breakout.Column).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                            ElseIf IsEmpty(cells(cell.row, or_breakout.Column).Offset(0, 1)) Then
                                cells(cell.row, or_breakout.Column).Offset(0, 1).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                            Else
                                cells(cell.row, or_breakout.Column).End(xlToRight).Offset(0, 1).Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                            End If
                        Else
                            no_cat_counter = no_cat_counter + 1
                            If no_cat_counter >= categories.Columns.Count Then
                                If IsEmpty(cells(cell.row, or_breakout.Column)) Then
                                    cells(cell.row, or_breakout.Column).Value = cat & Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                                ElseIf IsEmpty(cells(cell.row, or_breakout.Column).Offset(0, 1)) Then
                                    cells(cell.row, or_breakout.Column).Offset(0, 1).Value = cat & Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                                Else
                                    cells(cell.row, or_breakout.Column).End(xlToRight).Offset(0, 1).Value = cat & Right(cell.Value, Len(cell.Value) - InStr(cell.Value, " OR ") - 3)
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next
        cell.ClearContents
    End If
Next

Range(or_breakout, or_breakout.Offset(0, 10)).FormulaR1C1 = "=COUNTIFS(R" & top_row_num + 1 & "C:R10000C,""<>""&"""")"

For Each cell In Range(or_breakout, or_breakout.Offset(0, 10))
    If cell.Value = 0 Then
        cell.ClearContents
    End If
Next




End Function
Function OrCleanUp(top_row_num, or_breakout, rule_range, rule_one)

Dim or_col As Range
Dim or_col_count As Integer
Dim cell As Range
Dim num As Integer

Application.Calculation = xlCalculationAutomatic

If IsEmpty(or_breakout.Offset(0, 1)) Then
    Set or_col = or_breakout
    or_col_count = or_breakout.Columns.Count
Else
    Set or_col = Range(or_breakout, or_breakout.End(xlToRight))
    or_col_count = or_col.Columns.Count
End If

num = 1

For Each cell In or_col.Offset(1, 0)
    cell.Value = "Or Rule " & num
    num = num + 1
Next

num = 1
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
    Range(cells(top_row_num, 1), cells(top_row_num, or_breakout.Column)).AutoFilter Field:=Rows(top_row_num).Find("Or Rule 1", LookAt:=xlWhole).Column, Criteria1:="<>"
Else
    Range(cells(top_row_num, 1), cells(top_row_num, or_breakout.Column)).AutoFilter Field:=Rows(top_row_num).Find("Or Rule 1", LookAt:=xlWhole).Column, Criteria1:="<>"
End If

While num <> or_col_count
    Sheets("Data Audit").Activate
    Range(cells(top_row_num, or_breakout.Column), cells(top_row_num, or_breakout.Column).End(xlDown).Offset(0, num - 1)).EntireColumn.Hidden = True
    Range(cells(top_row_num, 1), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, or_breakout.Column).Offset(0, num).Column)).SpecialCells(xlCellTypeVisible).Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Or Sheet " & num
    Range("A2").PasteSpecial xlPasteValues
    num = num + 1
Wend

Sheets("Data Audit").Activate

cells.EntireColumn.Hidden = False
ActiveSheet.ShowAllData

Range(or_breakout.Offset(0, 1), or_breakout.Offset(0, 1).End(xlToRight)).EntireColumn.Delete Shift:=xlToLeft

num = 1

While num <> or_col_count
    Sheets("Or Sheet " & num).Activate
    Range("A2", cells(2, Rows(2).Find("Or Rule ", LookAt:=xlPart).Column)).AutoFilter Field:=Rows(2).Find("Or Rule ", LookAt:=xlPart).Column, Criteria1:="<>"
    If IsEmpty(Range("A2").Offset(1, 0)) = False Then
        Range(Range("A2").Offset(1, 0), cells(2, Rows(2).Find("Or Rule ", LookAt:=xlPart).Column).End(xlDown)).Copy
    End If
    Sheets("Data Audit").Activate
    cells(top_row_num, 1).End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
    num = num + 1
Wend

Sheets("Data Audit").Activate

Application.DisplayAlerts = False
num = 1

While num <> or_col_count
    Sheets("Or Sheet " & num).Delete
    num = num + 1
Wend

Sheets("Data Audit").Activate

Application.DisplayAlerts = True

Set rule_range = Range(rule_one.Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, rule_one.End(xlToRight).End(xlToRight).Column))

rule_range.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlToLeft


For Each cell In Range(cells(top_row_num + 1, rule_one.Column), cells(cells(top_row_num, 1).End(xlDown).row, rule_one.Column))
    If IsEmpty(cell) Then
        cell.EntireRow.Delete Shift:=xlUp
    End If
Next

    


cells(1, Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole).Column).Copy
For Each cell In Range("AQ1", Range("AQ1").Offset(0, 15))
    cell.PasteSpecial xlPasteFormulas
Next

For Each cell In Range(cells(1, rule_one.Column), cells(1, rule_one.Column).Offset(0, 15))
    If cell.Value = 0 Then
        cell.ClearContents
    End If
Next



Range(rule_one, rule_one.Offset(0, 50)).ClearContents

Set rule_row = Range(rule_one, cells(1, rule_one.Column).End(xlToRight).Offset(1, 0))

num = 1
For Each cell In rule_row
    cell.Value = "Rule " & num
    num = num + 1
Next


End Function

Function MoveCategories(rule_range, categories)

Dim cell As Range
Dim n As Range
Dim counter As Integer
Dim str As String
Dim start_num As Long
Dim length As Integer

Application.Calculation = xlCalculationAutomatic

For Each cell In rule_range
    For Each n In categories
        If InStr(cell.Text, n) Then
            str = Right(cell.Text, Len(cell.Value) - Len(n.Value))
            If IsEmpty(cells(cell.row, n.Column)) Then
                cells(cell.row, n.Column).Value = str
            Else
                cells(cell.row, n.Column).Value = cells(cell.row, n.Column).Value & " AND " & str
            End If
        End If
    Next
Next

End Function

Function overlap_code(top_row_num, rule_one)

Dim categories2 As Range
Dim cell As Range
Dim num As Integer
Dim gp_top As Range
Dim cg_top As Range
Dim grade_profiles As Range
Dim copied_gps As Range
Dim copied_cgs As Range
Dim bottom_row As Long
Dim copied_criteria As Range
Dim copied_criteria_top As Range
Dim matches_rng As Range
Dim match_cell As Range
Dim match_rng As Range
Dim cat2_start As Range
Dim criteria_check As Range
Dim criteria_cehck2 As Range
Dim criteria_check_top As Range
Dim active_criteria_top As Range
Dim active_criteria As Range
Dim equals_criteria_top As Range
Dim equals_criteria As Range
Dim new_profiles As Range
Dim new_profiles_bottom As Range
Dim macro_formulas As Range
Dim macro_formulas2 As Range
Dim macro_cell As Range
Dim cell_string As String
Dim rule_array As New Collection
Dim adding As String
Dim x As Integer
Dim item As Variant
Dim new_profile As Range









Set categories2 = Range(Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(top_row_num).Find("Time Type <> ", LookAt:=xlWhole))
bottom_row = cells(top_row_num, 1).End(xlDown).row



Set cat2_start = rule_one.End(xlToRight).Offset(0, 5)

Range(Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(top_row_num).Find("Time Type <> ", LookAt:=xlWhole)).Copy
num = 1
While num <> 3
    If num = 1 Then
        cat2_start.PasteSpecial xlPasteValues
    Else
        cat2_start.End(xlToRight).Offset(0, 1).PasteSpecial xlPasteValues
    End If
    num = num + 1
Wend


cat2_start.Offset(0, -2).Value = "Compensation Grade Copied"
cat2_start.Offset(0, -1).Value = "Grade Profile Copied"
cat2_start.End(xlToRight).Offset(0, 1).Value = "Match?"
cat2_start.End(xlToRight).Offset(0, 1).Value = "Grade Profile Match Copy"

Set gp_top = Rows(top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole)
Set cg_top = Rows(top_row_num).Find("Compensation Grade", LookAt:=xlWhole)

If IsEmpty(gp_top.Offset(1, 0)) = False Then
    Set grade_profiles = Range(gp_top.Offset(1, 0), gp_top.End(xlDown))
End If

'right here

Set copied_gps = Range(Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole).Offset(1, 0), cells(bottom_row, Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole).Column))
Set copied_cgs = Range(Rows(top_row_num).Find("Compensation Grade Copied", LookAt:=xlWhole).Offset(1, 0), cells(bottom_row, Rows(top_row_num).Find("Compensation Grade Copied", LookAt:=xlWhole).Column))

'this is the whole range for the copying of each grade profile rules
Set copied_criteria = Range( _
        Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole).Offset(1, 1), _
        cells(bottom_row, Range(Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole), _
        Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("Time Type <> ", LookAt:=xlWhole).Column))

'this sets the range for the top row of the data in the grade profile that we are iterating
Set copied_criteria_top = Range(Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole).Offset(0, 1), Range(Rows(top_row_num).Find("Grade Profile Copied"), Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("Time Type <> ", LookAt:=xlWhole))



'copied_criteria_top.Select
'MsgBox "Here's copied_criteria_top"
'copied_criteria_top.Item(copied_criteria_top.Columns.Count).Offset(0, 1).Select
'MsgBox "Here's copied_criteria_top column.count.offset(0,1)"

'this sets the range for the top row of formulas to check, will need a second one
Set criteria_check = Range( _
        Range(copied_criteria_top.item(copied_criteria_top.Columns.Count), Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("Certification = ", LookAt:=xlWhole), _
        Range(copied_criteria_top.item(copied_criteria_top.Columns.Count).Offset(0, 1), Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("Time Type <> ", LookAt:=xlWhole))
        
'criteria_check.Select
'MsgBox "Here's criteria_check"
        
        
'Set criteria_check2 = Range( _
'        Range(criteria_check.Item(criteria_check.Columns.Count), Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("Certification = ", LookAt:=xlWhole), _
'        Range(criteria_check.Item(criteria_check.Columns.Count).Offset(0, 1), Rows(top_row_num).Find("Match?", LookAt:=xlWhole)).Find("TM Sup Org Location <> ", LookAt:=xlWhole))

For Each cell In criteria_check.Offset(1, 0)
    Set criteria_check_top = cells(top_row_num, cell.Column)
    Set active_criteria_top = Rows(top_row_num).Find(criteria_check_top.Value, LookAt:=xlWhole)
    Set active_criteria = cells(cell.row, active_criteria_top.Column)
    'If InStr(1, criteria_check_top.Value, " <> ") Or InStr(1, criteria_check_top.Value, " >= ") Then
    
    'equals_criteria is for the rules that we are copying per grade profile/team member
    
    Set equals_criteria_top = copied_criteria_top.Find(active_criteria_top.Value, LookAt:=xlWhole)
    'Else
    '   Set equals_criteria_top = Rows(top_row_num).Find(Left(active_criteria_top.Value, Len(active_criteria_top.Value) - 3), LookAt:=xlWhole)
    'End If
    Set equals_criteria = cells(cell.row, equals_criteria_top.Column)
       
    
    'This part puts the formulas in to check if criteria is being met
    
    If InStr(1, criteria_check_top.Value, " = ") Then
        cell.FormulaR1C1 = "=IF(ISBLANK(RC" & active_criteria.Column & "),IF(ISBLANK(RC" & active_criteria.Offset(0, 10).Column & _
        "),"""",IF(AND(ISBLANK(RC" & active_criteria.Offset(0, 10).Column & ")=FALSE,ISBLANK(RC" & equals_criteria.Column & ")),"""",IF(ISNUMBER(SEARCH(RC" & equals_criteria.Column & ",RC" & _
        active_criteria.Offset(0, 10).Column & ")),""No"",""Yes""))),IF(AND(ISBLANK(RC" & active_criteria.Column & ")=FALSE,ISBLANK(RC" & equals_criteria.Column & ")),"""",IF(RC" & active_criteria.Column & "=RC" & _
        equals_criteria.Column & ",""Yes"",""No"")))"
    ElseIf InStr(1, criteria_check_top.Value, " >= ") Then
        cell.FormulaR1C1 = "=IF(OR(ISBLANK(RC" & active_criteria.Column & "),AND(ISBLANK(RC" & active_criteria.Column & ")=FALSE, ISBLANK(RC" & equals_criteria.Column & "))),"""",IF(RC" & equals_criteria.Column & ">=RC" & active_criteria.Column & ",""Yes"",""No""))"
    ElseIf InStr(1, criteria_check_top.Value, " <> ") Then
        cell.FormulaR1C1 = "=IF(ISBLANK(RC" & active_criteria.Offset(0, -10).Column & "),"""",IF(AND(ISBLANK(RC" & active_criteria.Offset(0, -10).Column & ")=FALSE,ISBLANK(RC" & equals_criteria.Column & _
        ")),"""",IF(ISNUMBER(SEARCH(RC" & active_criteria.Offset(0, -10).Column & ",RC" & equals_criteria.Column & ")),""No"",""Yes"")))"
    ElseIf InStr(1, criteria_check_top.Value, " < ") Then
        cell.FormulaR1C1 = "=IF(OR(ISBLANK(RC" & active_criteria.Column & "),AND(ISBLANK(RC" & active_criteria.Column & ")=FALSE, ISBLANK(RC" & equals_criteria.Column & "))),"""",IF(RC" & equals_criteria.Column & "<RC" & active_criteria.Column & ",""Yes"",""No""))"
    End If
    
Next
        


'Match?   BB is Compensation Grade Copied, A is First Column, obviously, BC is Grade Profile Copied,
'L is that rows Grade Profile, BN is the copied Certification = through TM Sup Org Location <>

Rows(top_row_num).Find("Match?", LookAt:=xlWhole).Offset(1, 0).FormulaR1C1 = "=IF(AND(RC" & Rows(top_row_num).Find("Compensation Grade Copied", LookAt:=xlWhole).Column & _
        "=RC1,RC" & Rows(top_row_num).Find("Grade Profile Copied", LookAt:=xlWhole).Column & "<>RC" & Rows(top_row_num).Find("Compensation Grade Profile", LookAt:=xlWhole).Column & _
        ",COUNTIFS(RC" & criteria_check.Find("Certification = ", LookAt:=xlWhole).Column & ":RC" & criteria_check.Find("Time Type <> ", LookAt:=xlWhole).Column & _
        ",""No"")=0,COUNTIFS(RC" & criteria_check.Find("Certification = ", LookAt:=xlWhole).Column & ":RC" & criteria_check.Find("Time Type <> ", LookAt:=xlWhole).Column & _
        ",""Yes"")>=0),""Yes"",""No"")"
grade_profiles.Copy
'Grade Profile Match Copy
Rows(top_row_num).Find("Grade Profile Match Copy", LookAt:=xlWhole).Offset(1, 0).PasteSpecial xlPasteValues


'change this to include match formula column

Set matches_rng = Range(criteria_check.Find("Certification = ", LookAt:=xlWhole).Offset(1, 0), cells(bottom_row, Rows(top_row_num).Find("Match?", LookAt:=xlWhole).Column))
Range(matches_rng.item(1), matches_rng.item(1).Offset(0, matches_rng.Columns.Count - 1)).Copy

matches_rng.PasteSpecial xlPasteFormulas

Sheets.Add After:=Sheets("Data Audit")
ActiveSheet.Name = "Matches"

Sheets("Data Audit").Activate
Range(Rows(top_row_num).Find("Compensation Grade Copied", LookAt:=xlWhole), Rows(top_row_num).Find("Grade Profile Match Copy", LookAt:=xlWhole)).Copy
Sheets("Matches").Activate
Range("A2").PasteSpecial xlPasteValues

Sheets("Data Audit").Activate

Set match_rng = Range(Rows(top_row_num).Find("Match?", LookAt:=xlWhole).Offset(1, 0), Rows(top_row_num).Find("Match?", LookAt:=xlWhole).End(xlDown))


On Error Resume Next
    Sheets("New Grade Profiles").Activate
If Err.Number = 0 Then
    If IsEmpty(Range("D4")) Then
        Set new_profiles = Sheets("New Grade Profiles").Range("D3")
    Else
        Set new_profiles = Sheets("New Grade Profiles").Range("D3", Range("C2").End(xlDown).Offset(0, 1))
    End If
End If
Err.Clear
On Error GoTo 0

'Set macro_formulas = Range(criteria_check.Find("Certification <> ", LookAt:=xlWhole).Offset(1, 0), Cells(bottom_row, criteria_check.Find("Loan Channel <> ", LookAt:=xlWhole).Column))
'Set macro_formulas2 = Range(criteria_check.Find("TM Location <> ", LookAt:=xlWhole).Offset(1, 0), Cells(bottom_row, criteria_check.Find("TM Sup Org Location <> ", LookAt:=xlWhole).Column))

For Each cell In grade_profiles
    On Error Resume Next
        Sheets("New Grade Profiles").Activate
    If Err.Number = 0 Then
        For Each new_profile In new_profiles
            If cell.Value = new_profile.Value Then
                Sheets("Data Audit").Activate
                copied_gps = cell.Value
                copied_cgs = cells(cell.row, cg_top.Column)
                Range(cells(cell.row, Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole).Column), cells(cell.row, Rows(top_row_num).Find("Time Type <> ", LookAt:=xlWhole).Column)).Copy
                Application.DisplayAlerts = False
                copied_criteria.PasteSpecial xlPasteValues
                Application.DisplayAlerts = True
                ActiveSheet.Calculate
                num = 0
                For Each match_cell In match_rng
                    If InStr(1, match_cell.Value, "Yes") Then
                        num = num + 1
                    End If
                Next
                If num > 0 Then
                    If ActiveSheet.AutoFilterMode = True Then
                        ActiveSheet.AutoFilterMode = False
                    End If
                    Range(cells(top_row_num, 1), Rows(top_row_num).Find("Grade Profile Match Copy", LookAt:=xlWhole)).AutoFilter Field:=Rows(top_row_num).Find("Match?", LookAt:=xlWhole).Column, Criteria1:="Yes"
                    'MsgBox "Hey", vbOKCancel
                    Range(Rows(top_row_num).Find("Compensation Grade Copied", LookAt:=xlWhole).Offset(1, 0), Rows(top_row_num).Find("Grade Profile Match Copy", LookAt:=xlWhole).End(xlDown)).Copy
                    Sheets("Matches").Activate
                    If IsEmpty(Range("A3")) = True Then
                       Range("A3").PasteSpecial xlPasteValues
                    Else
                        Range("A2").End(xlDown).Offset(1, 0).PasteSpecial xlPasteValues
                    End If
                    Sheets("Data Audit").Activate
                    ActiveSheet.ShowAllData
                End If
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
Next


Sheets("Matches").Activate

Range("A2").End(xlToRight).Offset(0, 1).Value = "Dupes?"
Range(Range("A2").End(xlToRight).Offset(1, 0), Range("A2").End(xlToRight).Offset(0, -1).End(xlDown).Offset(0, 1)).FormulaR1C1 = "=RC1&RC2&RC[-1]"

Range(Range("A2").Offset(1, 0), Range("A2").End(xlToRight).End(xlDown)).RemoveDuplicates Columns:=45, Header:=xlNo

Range("A2").End(xlToRight).EntireColumn.Delete Shift:=xlToLeft


Range("C1:AQ1").EntireColumn.Hidden = True

End Function

Sub Overlap()


Dim top_row_num As Long
Dim init_rule_top As Range
Dim init_rule_range As Range
Dim sub_rule_range As Range
Dim rule_range As Range
Dim or_breakout As Range
Dim rule_one As Range
Dim cell As Range
Dim counter As Integer
Dim categories As Range

On Error Resume Next
    Sheets("Data Audit").Activate
If Err.Number = 9 Then
    MsgBox "Please Update the Sheet Name"
    ActiveSheet.Name = InputBox("what should the name be?")
    Sheets("Data Audit").Activate
End If
Err.Clear
On Error GoTo 0

Sheets("Data Audit").Activate
top_row_num = Columns(1).Find("Compensation Grade", LookAt:=xlWhole).row

ClearOut top_row_num

Sheets("Data Audit").Activate
Set init_rule_top = Rows(top_row_num).Find("Gr Profile: Eligibility Rules", LookAt:=xlWhole)
Set init_rule_range = Range(init_rule_top.Offset(1, 0), init_rule_top.End(xlDown))

FixDataIssues init_rule_range

Set sub_rule_range = init_rule_range.Offset(0, 1)
FirstOrLogic top_row_num, sub_rule_range, init_rule_range
AndLogic top_row_num

Set rule_range = Range(Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole).Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column))

Set or_breakout = rule_range.End(xlUp).End(xlToRight).Offset(0, 5)

Set rule_one = Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole)

counter = 1
While counter <> 0
   counter = 0
    For Each cell In rule_range
        If InStr(1, cell.Value, " OR ") Then
            counter = counter + 1
        End If
    Next
    If counter > 0 Then
        SecondOrLogic top_row_num, rule_range, or_breakout
        OrCleanUp top_row_num, or_breakout, rule_range, rule_one
    End If
Wend
            
Set rule_range = Range(Rows(top_row_num).Find("Rule 1", LookAt:=xlWhole).Offset(1, 0), cells(cells(top_row_num, 1).End(xlDown).row, cells(top_row_num, 1).End(xlToRight).Column))
        
FixDataIssues rule_range

Set categories = Range(Rows(top_row_num).Find("Certification = ", LookAt:=xlWhole), Rows(top_row_num).Find("Time Type <> ", LookAt:=xlWhole))

MoveCategories rule_range, categories

overlap_code top_row_num, rule_one


End Sub





Sub Wrangle()

Dim cell As Range
Dim cell2 As Range
Dim gradeRng As Range
Dim gradeProfileRng As Range
Dim commissionPlanRng As Range
Dim unassComponentRng As Range
Dim assIneligComponentRng As Range
Dim assEligComponentRng As Range
Dim topRow As Integer
Dim newFields As Variant
Dim dataSheet As String
Dim colNums As Variant
Dim firstNewCol As Integer
Dim bottomRow As Integer
Dim startFormRng As Range


Dim EliGradeCol As Integer
Dim EliGradeProfCol As Integer
Dim EliCommPlanCol As Integer

Dim IneligGradeCol As Integer
Dim IneligGradeProfileCol As Integer
Dim IneligCommPlanCol As Integer

Dim eligGrades As Range
Dim eligGradeProfiles As Range
Dim ineligGrade As Range
Dim ineligGradeProfiles As Range


Dim startTime As Double


Dim formRng As Range
Dim formulas As Range



Dim counter As Integer

startTime = Timer

'NEED TO SETUP THE SHEET

'TWO SHEETS: Wrangled = Raw Data, References = 3 columns of Grade, Grade Profiles, and Commission Plans starting at row 5, column b

'will need a Data Audit for Grade and Grade Profiles, and TM List for the Commission Plans

'Remove the dupes and put them in columns b, c, and d


dataSheet = "TM Comp Audit"
    
Sheets(dataSheet).Activate
topRow = Columns(1).Find("Team Member", LookAt:=xlWhole).row
bottomRow = cells(topRow, 1).End(xlDown).row
firstNewCol = Rows(topRow).Find("Unassigned Eligible Compensation Components").Offset(0, 1).Column

'newFields = Array("Ass. Inelig. Grade", "Ass. Inelig. Grade Profile", "Ass. Inelig. Commission Plan", "Unass. Elig. Grade", "Unass. Elig. Grade Profile", "Unass. Elig. Commission Plan")

'For i = 0 To UBound(newFields)
'    cells(topRow, 1).End(xlToRight).Offset(0, 1).Value = newFields(i)
'Next i

EligGradeCol = firstNewCol
EligGradeProfCol = Rows(topRow).Find("Elig Grade Profile 1").Column
EligCommPlanCol = Rows(topRow).Find("Elig. Comm Plan 1").Column

IneligGradeCol = Rows(topRow).Find("Inelig Grade 1").Column
IneligGradeProfileCol = Rows(topRow).Find("Inelig Grade Profile 1").Column
IneligCommPlanCol = Rows(topRow).Find("Inelig Comm Plan 1").Column


'cells(topRow, 1).Copy
'Range(cells(topRow, 1), cells(topRow, 1).End(xlToRight)).Select
'With Selection
'    .PasteSpecial xlPasteFormats
'    .ColumnWidth = 20
'    .Columns.AutoFit
'End With

Range("A1").Select

Application.CutCopyMode = False

Sheets("References").Activate

'Eventually will automatically update with new data

'Need to Automatically Create these references
'This absolute references are ok here because that information stays in the same place

Set gradeRng = Range("B20", Range("B20").End(xlDown))
Set gradeProfileRng = Range("C20", Range("C20").End(xlDown))

Sheets("Compensation Plans").Activate

Set commissionPlanRng = Range("A3", Range("A3").End(xlDown))

Sheets(dataSheet).Activate

Set assEligComponentRng = Range("G10", cells(cells(topRow, 1).End(xlDown).row, Range("G10").Column))
Set assIneligComponentRng = Range("H10", cells(cells(topRow, 1).End(xlDown).row, Range("H10").Column))
Set unassComponentRng = Range("I10", cells(cells(topRow, 1).End(xlDown).row, Range("I10").Column))




    
pullOutData gradeRng, assIneligComponentRng, IneligGradeCol, "No"
pullOutData gradeRng, unassComponentRng, EligGradeCol, "No"
pullOutData gradeRng, assEligComponentRng, EligGradeCol, "No"

pullOutData gradeProfileRng, assIneligComponentRng, IneligGradeProfileCol, "No"
pullOutData gradeProfileRng, unassComponentRng, EligGradeProfCol, "Yes"
pullOutData gradeProfileRng, assEligComponentRng, EligGradeProfCol, "Yes"

pullOutData commissionPlanRng, assIneligComponentRng, IneligCommPlanCol, "No"
pullOutData commissionPlanRng, unassComponentRng, EligCommPlanCol, "No"
pullOutData commissionPlanRng, assEligComponentRng, EligCommPlanCol, "No"

Application.Calculate




Set startFormRng = Rows(9).Find("Elig Grade Profile 1 Check for in Grade").Offset(-1, 0)

Set formRng = Range(startFormRng.Offset(2, 0), cells(bottomRow, startFormRng.Offset(1, 0).End(xlToRight).Column))
Set formulas = Range(startFormRng, startFormRng.End(xlToRight))
formulas.Copy
formRng.PasteSpecial xlPasteFormulas

ActiveSheet.Calculate

removeIncorrect
    
ActiveSheet.Calculate

'Copy Eligible Grades over to end and remove blanks, then copy back over the original info
'NEED TO UPDATE THESE TO BE DYNAMIC

Set eligGrades = Range(Rows(topRow).Find("Elig Grade 1", LookAt:=xlWhole).Offset(1, 0), cells(bottomRow, Rows(topRow).Find("Elig Grade 5", LookAt:=xlWhole).Column))
Set eligGradeProfiles = Range(Rows(topRow).Find("Elig Grade Profile 1", LookAt:=xlWhole).Offset(1, 0), cells(bottomRow, Rows(topRow).Find("Elig Grade Profile 10", LookAt:=xlWhole).Column))

Set ineligGrades = Range(Rows(topRow).Find("Inelig Grade 1", LookAt:=xlWhole).Offset(1, 0), cells(bottomRow, Rows(topRow).Find("Inelig Grade 5", LookAt:=xlWhole).Column))
Set ineligGradeProfiles = Range(Rows(topRow).Find("Inelig Grade Profile 1", LookAt:=xlWhole).Offset(1, 0), cells(bottomRow, Rows(topRow).Find("Inelig Grade Profile 10", LookAt:=xlWhole).Column))

fillInToLeft eligGrades, bottomRow, "No"
fillInToLeft eligGradeProfiles, bottomRow, "Yes"
fillInToLeft ineligGrades, bottomRow, "No"
fillInToLeft ineligGradeProfiles, bottomRow, "Yes"

Application.Calculate

copyToPreEIB

MsgBox "The macro has finished in " & Round(Timer - startTime, 2) & " seconds or " & Round((Timer - startTime) / 60, 2) & " minutes."


End Sub


Function pullOutData(refRng, dataRng, destinationCol, isEligGradeProfile)

Dim cell As Range
Dim cell2 As Range


For Each cell In refRng
    For Each cell2 In dataRng
        If InStr(1, cell2.Value, cell.Value) Then
            If IsEmpty(cells(cell2.row, destinationCol)) Then
                cells(cell2.row, destinationCol).Value = cell.Value
            ElseIf IsEmpty(cells(cell2.row, destinationCol).Offset(0, 1)) Then
                cells(cell2.row, destinationCol).Offset(0, 1).Value = cell.Value
            Else
                cells(cell2.row, destinationCol).End(xlToRight).Offset(0, 1).Value = cell.Value
            End If
        
        End If
    Next
Next
        

End Function


Function removeIncorrect()

Dim bottomRow As Integer
Dim firstCol As Integer
Dim lastCol As Integer
Dim checkCol As Integer

'NEED TO FIX THESE TO BE DYNAMIC

firstCol = Rows(9).Find("Elig Grade Profile 1", LookAt:=xlWhole).Column
lastCol = Rows(9).Find("Elig Grade Profile 10", LookAt:=xlWhole).Column


bottomRow = Range("A10").End(xlDown).row

'This is for eligible grade profiles


For Each cell In Range(cells(10, firstCol), cells(bottomRow, lastCol))
    If IsEmpty(cell) = False Then
        If cell.Offset(0, 35).Value < 1 Then
            cell.ClearContents
        End If
    End If
Next


'this is for eligible grades
'NEED TO UPDATE THIS TO BE DYNAMIC


firstCol = Rows(9).Find("Elig Grade 1", LookAt:=xlWhole).Column
lastCol = Rows(9).Find("Elig Grade 5", LookAt:=xlWhole).Column

checkCol = Rows(9).Find("Grade Based on Job Profile", LookAt:=xlWhole).Column

For Each cell In Range(cells(10, firstCol), cells(bottomRow, lastCol))
    If IsEmpty(cell) = False Then
        If cell.Value <> cells(cell.row, checkCol).Value Then
            cell.ClearContents
        End If
    End If
Next

'this is for ineligible grade profiles
'NEED TO UPDATE THIS TO BE DYNAMIC

firstCol = Rows(9).Find("Inelig Grade Profile 1", LookAt:=xlWhole).Column
lastCol = Rows(9).Find("Inelig Grade Profile 10", LookAt:=xlWhole).Column


For Each cell In Range(cells(10, firstCol), cells(bottomRow, lastCol))
    If IsEmpty(cell) = False Then
        If cell.Offset(0, 25).Value < 1 Then
            cell.ClearContents
        End If
    End If
Next
        
'this is for ineligible grades
'NEED TO UPDATE THIS TO BE DYNAMIC

firstCol = Rows(9).Find("Inelig Grade 1", LookAt:=xlWhole).Column
lastCol = Rows(9).Find("Inelig Grade 5", LookAt:=xlWhole).Column

checkCol = Rows(9).Find("Grade Based on Job Profile", LookAt:=xlWhole).Column

For Each cell In Range(cells(10, firstCol), cells(bottomRow, lastCol))
    If IsEmpty(cell) = False Then
        If cell.Value <> cells(cell.row, checkCol).Value Then
            cell.ClearContents
        End If
    End If
Next

End Function


Function fillInToLeft(rng, bottomRow, isGradeProfile)

Dim newRng As Range
Dim cell As Range

If isGradeProfile = "Yes" Then
    Set newRng = Range(Range("A9").End(xlToRight).Offset(1, 1), cells(bottomRow, Range("A9").End(xlToRight).Offset(0, 10).Column))
Else
    Set newRng = Range(Range("A9").End(xlToRight).Offset(1, 1), cells(bottomRow, Range("A9").End(xlToRight).Offset(0, 5).Column))
End If

For Each cell In rng
    If cell.Value = "" Then
        cell.ClearContents
    End If
Next

rng.Copy
newRng.PasteSpecial xlPasteValues

newRng.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlToLeft

newRng.Copy
rng.PasteSpecial xlPasteValues

newRng.ClearContents

End Function

Function copyToPreEIB()

Dim topRow As Range
Dim bottomRowNum As Integer


Sheets("TM Comp Audit").Activate

Set topRow = Range("A9", Range("A9").End(xlToRight))

Application.Calculate

If Range("CE4").Value > 0 Then
    If ActiveSheet.AutoFilterMode = True Then
        ActiveSheet.AutoFilterMode = False
    End If
    topRow.AutoFilter Field:=topRow.Find("Needs Something Updated?").Column, Criteria1:="Yes"
    Range(Range("B9").Offset(1, 0), Range("B9").End(xlDown)).SpecialCells(xlCellTypeVisible).Copy
    Sheets("Pre EIB").Activate
    Range("C5").PasteSpecial xlPasteValues
    bottomRowNum = Range("C5").End(xlDown).row
    Range("B3").Copy
    Range("B5", cells(bottomRowNum, Range("A5").Column)).PasteSpecial xlPasteFormulas
    Range("B5", cells(bottomRowNum, Range("A5").Column)).PasteSpecial xlPasteFormats
    Range("E3:BY3").Copy
    Range("E5", cells(bottomRowNum, Range("BY4").Column)).PasteSpecial xlPasteFormulas
    Range("E5", cells(bottomRowNum, Range("BY4").Column)).PasteSpecial xlPasteFormats
    ActiveSheet.Calculate
 
    
    
    
End If



End Function
Attribute VB_Name = "AllSchedulesMod"
Option Explicit
' read 3W, 8P and 3P schedules and copy into All Schedules
Public Sub readSchedules()
    Dim initials As String
    Dim initialsArray As Object
    Dim counter As Integer
    Dim schedCell As range
    Dim schedCell2 As range
    Dim schedCell3 As range
    Dim initialsBox As range
    Dim therCounter As Integer
    Dim r1 As range
    Dim r2 As range
    Dim r3 As range
    Dim cell1 As range
    Dim cell2 As range
    Dim str1 As String
    Dim str2 As String
    Dim i As Integer
         
    
    ' create initials arraylist
    Set initialsArray = CreateObject("System.Collections.ArrayList")
    
    ' loop through 3W to get initials; add to arraylist
    For Each schedCell In Sheets("3W Schedule").range("SchedGrid3W")
        ' check for empty cell, empty string, numbers & TMG procedure
        If Not IsEmpty(schedCell) And schedCell.value <> " " And schedCell.value <> "" And Not IsNumeric(Left(schedCell.value, 1)) And Trim(schedCell.value) <> "TMG" Then
            initials = schedCell.value
            If isGray(initials) = False And LCase(Trim(initials)) <> "lunch" Then
                 initials = returnInitials(initials)
                If Not initialsArray.Contains(initials) And initials <> "NOTE" Then
                    initialsArray.Add (UCase(initials))
                End If
            End If
        End If
    Next schedCell
    
    ' loop through 8P to get initials; add to arraylist
    For Each schedCell2 In Sheets("8P Schedule").range("SchedGrid8P")
        ' check for empty cell, empty string, numbers & TMG procedure
        If Not IsEmpty(schedCell2) And schedCell2.value <> " " And schedCell2.value <> "" And Not IsNumeric(Left(schedCell2.value, 1)) And Trim(schedCell2.value) <> "TMG" Then
            initials = schedCell2.value
            If isGray(initials) = False And LCase(Trim(initials)) <> "lunch" Then
                 initials = returnInitials(initials)
                If Not initialsArray.Contains(initials) And initials <> "NOTE" Then
                    initialsArray.Add (UCase(initials))
                End If
            End If
        End If
    Next schedCell2
    
    ' loop through 3P to get initials; add to arraylist
    For Each schedCell3 In Sheets("3P Schedule").range("SchedGrid3P")
        ' check for empty cell, empty string, numbers & TMG procedure
        If Not IsEmpty(schedCell3) And schedCell3.value <> " " And schedCell3 <> "" And Not IsNumeric(Left(schedCell3.value, 1)) And Trim(schedCell3.value) <> "TMG" Then
            initials = schedCell3.value
            If isGray(initials) = False And LCase(Trim(initials)) <> "lunch" Then
                 initials = returnInitials(initials)
                If Not initialsArray.Contains(initials) And initials <> "NOTE" Then
                    initialsArray.Add (UCase(initials))
                End If
            End If
        End If
    Next schedCell3
    
   ' add initials to All Schedules
   
   ' initialize counter
   counter = 0
   For Each initialsBox In Sheets("All Schedules").range("AllSchedTherapistInitialsBox")
        initialsBox.value = initialsArray.Item(counter)
        counter = counter + 1
        If counter > initialsArray.Count - 1 Then
            Exit For
        End If
   Next initialsBox
   
   ' look at each row in 3W, 8P & 3P schedules; if initials match initials in row
   ' copy room and schedule row and paste into all schedules
   ' initialize row counter
   i = 3
   For Each cell1 In Sheets("All Schedules").range("AllSchedTherapistInitialsBox")
    ' if the initials box is not empty, add a schedule
    If Not IsEmpty(cell1) And cell1.value <> "" And cell1.value <> " " Then
        ' look at 3W Schedule
        For Each r1 In Sheets("3W Schedule").range("SchedGrid3W").Rows
            ' "ADL EC" for instance
            str1 = "ADL " + CStr(cell1.value)
            ' "ADLEC" for instance; account for lack of space between ADL and initials
            str2 = "ADL" + CStr(cell1.value)
            ' if the row in 3W Schedule contains EC, ADL EC, or ADLEC, for instance
            If Application.WorksheetFunction.CountIf(r1, cell1.value) > 0 Or Application.WorksheetFunction.CountIf(r1, str1) > 0 Or Application.WorksheetFunction.CountIf(r1, str2) > 0 Then
                cell1.Offset(i, -14).value = r1.Cells(1, 1).value
                range(r1.Cells(1, 2), r1.Cells(1, 23)).Copy
                range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
                ' go to next row in schedule
                i = i + 1
            End If
        Next r1
        ' look at 8P schedule
        For Each r2 In Sheets("8P Schedule").range("SchedGrid8P").Rows
            str1 = "ADL " + CStr(cell1.value)
            str2 = "ADL" + CStr(cell1.value)
            If Application.WorksheetFunction.CountIf(r2, cell1.value) > 0 Or Application.WorksheetFunction.CountIf(r2, str1) > 0 Or Application.WorksheetFunction.CountIf(r2, str2) > 0 Then
                cell1.Offset(i, -14).value = r2.Cells(1, 1).value
                range(r2.Cells(1, 2), r2.Cells(1, 23)).Copy
                range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
                ' go to next row in schedule
                i = i + 1
            End If
        Next r2
        ' look at 3P schedule
        For Each r3 In Sheets("3P Schedule").range("SchedGrid3P").Rows
            str1 = "ADL " + CStr(cell1.value)
            str2 = "ADL" + CStr(cell1.value)
            If Application.WorksheetFunction.CountIf(r3, cell1.value) > 0 Or Application.WorksheetFunction.CountIf(r3, str1) > 0 Or Application.WorksheetFunction.CountIf(r3, str2) > 0 Then
                cell1.Offset(i, -14).value = r3.Cells(1, 1).value
                range(r3.Cells(1, 2), r3.Cells(1, 23)).Copy
                range(cell1.Offset(i, -10), cell1.Offset(i, 11)).PasteSpecial xlPasteValues
                ' go to next row in schedule
                i = i + 1
            End If
        Next r3
        ' reset row number for next schedule so you can start at the top
        i = 3
       
     End If

   Next cell1
   
   ' reset initialsArray
   Set initialsArray = Nothing
   ' get rid of cut copy mode
   Application.CutCopyMode = False
   
   
  

End Sub

' gets notes from All Therapists sheet and puts them into All Schedules sheet
Public Sub getNotes()
Dim cell As range
Dim initials As String
Dim note As String
Dim cell2 As range
Dim initialsBoxRng As range
Dim therapistsInitials As range
Dim allSchedNotes As range
Dim noteDict As Object
' create noted dictionary object
Set noteDict = CreateObject("Scripting.Dictionary")

Set therapistsInitials = Sheets("All Therapists").range("AllTherapistsInitials")
Set allSchedNotes = Sheets("All Schedules").range("AllSchedNoteCells")

' loop through all therapists to get initials and note; add to dictionary
' key = initials; value = note
    For Each cell In therapistsInitials
        If Not IsEmpty(cell) And cell.value <> "-" And cell.value <> " " Then
             initials = UCase(Trim(cell.value))
            note = cell.Offset(0, 25).value
            noteDict.Add Key:=initials, Item:=note
        End If
    Next cell
' put notes in note area on All Schedules
For Each cell2 In allSchedNotes
    cell2.value = noteDict(cell2.Offset(-25, 14).value)
Next cell2

' reset noteDict
Set noteDict = Nothing
End Sub
' clear highlighting and old data and populate schedules; scroll back to top of page
Public Sub createSchedules()
    ' declare variables
    Dim timeCreated As range
    
    ' turn off screen updating
   Application.ScreenUpdating = False
    
    ' clear highlighting and previous info
    Call clearAllSched
    ' set box for displaying time schedule was created
    Set timeCreated = Sheets("All Schedules").range("AllSchedulesTimeCreatedCell")
    
    ' disable alerts about replacing data in non-blank cells
    Application.DisplayAlerts = False
    
    ' read schedules
    Call readSchedules
    ' get notes
    Call getNotes
    ' show time created
    Call lastTimeCreated(timeCreated)
    ' apply highlighting
    Call schedCondFormat(Sheets("All Schedules").range("EvalKeyBox"), Sheets("All Schedules").range("IntKeyBox"), _
        Sheets("All Schedules").range("RoomsAllSchedules"), Sheets("All Schedules").range("AllSchedTherapistInitialsBox"), _
        Sheets("All Schedules"))
    ' go back to top of sheet
    ActiveWindow.ScrollRow = 1
    ' select cell at top of page to remove selection bar at bottom of file
    ActiveSheet.range("$AE$6").Select
    ' turn alerts back on
    Application.DisplayAlerts = True
    ' turn on screen updating
   Application.ScreenUpdating = True
    
End Sub



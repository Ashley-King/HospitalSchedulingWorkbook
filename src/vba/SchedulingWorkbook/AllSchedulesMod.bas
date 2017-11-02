Attribute VB_Name = "AllSchedulesMod"
Option Explicit

' module-scope variables
Private rowCounter As Integer
Private initialsArray As Object
' read 3W, 8P and 3P schedules and copy into All Schedules
Public Sub readSchedules()
    Dim counter As Integer
    Dim initialsBox As range
    Dim cell As range
    
    
    ' create initials arraylist
    Set initialsArray = CreateObject("System.Collections.ArrayList")
    Call createInitialsArray(Sheets("3W Schedule"), "SchedGrid3W")
    Call createInitialsArray(Sheets("8P Schedule"), "SchedGrid8P")
    Call createInitialsArray(Sheets("3P Schedule"), "SchedGrid3P")
   
    ' initialize counter
    counter = 0
    ' add initials to all schedules sheets
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
    rowCounter = 3
    For Each cell In Sheets("All Schedules").range("AllSchedTherapistInitialsBox")
        ' if the initials box is not empty, add a schedule
        If Not IsEmpty(cell) And cell.value <> "" And cell.value <> " " Then
            Call populateSchedules(Sheets("3W Schedule"), "SchedGrid3W", cell, rowCounter)
            Call populateSchedules(Sheets("8P Schedule"), "SchedGrid8P", cell, rowCounter)
            Call populateSchedules(Sheets("3P Schedule"), "SchedGrid3P", cell, rowCounter)
            ' reset row counter
            rowCounter = 3
        End If
    Next cell
   
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
Public Sub Create_All_Schedules()
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

Public Sub createInitialsArray(sheet As Worksheet, grid As String)
    Dim schedCell As range
    Dim initials As String
    
    ' loop through schedule to get initials; add to arraylist
    For Each schedCell In sheet.range(grid)
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
End Sub



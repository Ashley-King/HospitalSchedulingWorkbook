Attribute VB_Name = "AllDiddiesMod"
Option Explicit

' read 3W, 8P and 3P schedules and create diddies for each room that is occupied
Public Sub Create_All_Diddies()
    
    ' declare variables
    'Dim room3 As range

    Dim timeCreated As range

    ' turn screen updating off
    Application.ScreenUpdating = False
    
    ' clear previous data and highlighting
    Call clearDiddies
    
    ' set timeCreated box in diddies sheet
    Set timeCreated = Sheets("All Diddies").range("AllDiddiesTimeCreatedCell")
    
    ' disable alerts about replacing data in non-blank cells
    Application.DisplayAlerts = False
    
    ' loop through all diddies rooms; if =room does not have anyone listed in Y AND Z columns, skip
    ' it, otherwise add room to diddie; then add schedules
    ' 3W
    Call addRoomToDiddies("RoomsRow3WAllDiddies", Sheets("3W Schedule"), "Rooms3WSchedule")
    Call addScheduleToDiddies("RoomsRow3WAllDiddies", Sheets("3W Schedule"), "Rooms3WSchedule")
    ' 8P
    Call addRoomToDiddies("RoomsRow8PAllDiddies", Sheets("8P Schedule"), "Rooms8PSchedule")
    Call addScheduleToDiddies("RoomsRow8PAllDiddies", Sheets("8P Schedule"), "Rooms8PSchedule")
    ' 3P
    Call addRoomToDiddies("RoomsRow3PAllDiddies", Sheets("3P Schedule"), "Rooms3PSchedule")
    Call addScheduleToDiddies("RoomsRow3PAllDiddies", Sheets("3P Schedule"), "Rooms3PSchedule")
      
    ' show time the schedule was last created
    Call lastTimeCreated(timeCreated)
    ' go back to top of sheet
    ActiveWindow.ScrollRow = 1
    
    ' turn screen updating on
    Application.ScreenUpdating = True
    ' turn alerts back on
    Application.DisplayAlerts = True
    
End Sub

' loop through all diddies rooms; if room does not have anyone listed in Y AND Z columns, skip
' it, otherwise add room to diddie
Public Sub addRoomToDiddies(diddiesRooms As String, schedSheet As Worksheet, schedRooms As String)
    Dim roomCell As range
    Dim room As range
    
    For Each room In Sheets("All Diddies").range(diddiesRooms)
        For Each roomCell In schedSheet.range(schedRooms)
            If (roomCell.Offset(0, 23) <> " " And roomCell.Offset(0, 23) <> "") Or (roomCell.Offset(0, 24) <> " " _
                And roomCell.Offset(0, 24) <> "") Then
                If Sheets("All Diddies").range(diddiesRooms).Find(roomCell.value, LookAt:=xlWhole, searchorder:=xlByRows) Is Nothing Then
                    room.value = roomCell.value
                    Exit For
                End If
            End If
        Next roomCell
    Next room
End Sub

Public Sub addScheduleToDiddies(diddiesRooms As String, sched As Worksheet, schedRooms As String)
    Dim room As range
    Dim newRoom As range
    Dim newString As String
    Dim i As Integer
    Dim j As Integer
    
    ' add schedule to diddies
    For Each room In Sheets("All Diddies").range(diddiesRooms)
        Set newRoom = sched.range(schedRooms).Find(room.value, LookAt:=xlWhole, searchorder:=xlByRows)
        ' i will start two rows below room in the 6:30 time slot
        i = 2
        If Not newRoom Is Nothing Then
            For j = 1 To 22
            If Not IsEmpty(newRoom.Offset(0, j)) And Not newRoom.Offset(0, j) = " " And Not newRoom Is Nothing Then
                ' put value in getNameAndProf function and return name, lunch or gray
                newString = getNameAndProf(newRoom.Offset(0, j).value)
                ' if function returns therapist name, TMG or lunch, add to diddie
                If newString = "LUNCH" Or newString <> "GRAY" Then
                    room.Offset(i, 0).value = newString
                End If
            End If
            ' increment row on diddie schedule
            i = i + 1
        Next j
        End If
    Next room
End Sub

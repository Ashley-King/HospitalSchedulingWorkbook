Attribute VB_Name = "AllDiddiesMod"
Option Explicit

' read 3W, 8P and 3P schedules and create diddies for each room that is occupied
Public Sub Create_All_Diddies()
    
    ' declare variables
    Dim room As range
    Dim room2 As range
    Dim room3 As range
    Dim room4 As range
    Dim room5 As range
    Dim room6 As range
    Dim nameAndProf As Variant
    Dim i As Integer
    Dim j As Integer
    Dim roomCell As range
    Dim roomCell2 As range
    Dim roomCell3 As range
    Dim roomCell4 As range
    Dim newRoom As range
    Dim newString As String
    Dim timeCreated As range

    ' turn screen updating off
    Application.ScreenUpdating = False
    
    ' clear previous data and highlighting
    Call clearDiddies
    
    ' set timeCreated box in diddies sheet
    Set timeCreated = Sheets("All Diddies").range("AllDiddiesTimeCreatedCell")
    
    ' disable alerts about replacing data in non-blank cells
    Application.DisplayAlerts = False
    
     
    ' loop through 3W all diddies rooms; if 3W room does not have anyone listed in Y AND Z columns, skip
    ' it, otherwise add room to diddie
    For Each room In Sheets("All Diddies").range("RoomsRow3WAllDiddies")
        For Each roomCell In Sheets("3W Schedule").range("Rooms3WSchedule")
            If (roomCell.Offset(0, 23) <> " " And roomCell.Offset(0, 23) <> "") Or (roomCell.Offset(0, 24) <> " " _
                And roomCell.Offset(0, 24) <> "") Then
                If Sheets("All Diddies").range("RoomsRow3WAllDiddies").Find(roomCell.value, LookAt:=xlWhole, searchorder:=xlByRows) Is Nothing Then
                    room.value = roomCell.value
                    Exit For
                End If
            End If
        Next roomCell
    Next room
    
    
    ' loop through 8P all diddies rooms; if 8P room does not have anyone listed in Y AND Z columns, skip
    ' it, otherwise add room to diddie
    For Each room2 In Sheets("All Diddies").range("RoomsRow8PAllDiddies")
        For Each roomCell2 In Sheets("8P Schedule").range("Rooms8PSchedule")
            If (roomCell2.Offset(0, 23) <> " " And roomCell2.Offset(0, 23) <> "") Or (roomCell2.Offset(0, 24) <> " " _
            And roomCell2.Offset(0, 24) <> "") Then
                If Sheets("All Diddies").range("RoomsRow8PAllDiddies").Find(roomCell2.value, LookAt:=xlWhole, searchorder:=xlByRows) Is Nothing Then
                    room2.value = roomCell2.value
                    Exit For
                End If
            End If
        Next roomCell2
    Next room2
    
    ' loop through 3P all diddies rooms; if 8P room does not have anyone listed in Y AND Z columns, skip
    ' it, otherwise add room to diddie
    For Each room5 In Sheets("All Diddies").range("RoomsRow3PAllDiddies")
        For Each roomCell4 In Sheets("3P Schedule").range("Rooms3PSchedule")
            If (roomCell4.Offset(0, 23) <> " " And roomCell4.Offset(0, 23) <> "") Or (roomCell4.Offset(0, 24) <> " " _
            And roomCell4.Offset(0, 24) <> "") Then
                If Sheets("All Diddies").range("RoomsRow3PAllDiddies").Find(roomCell4.value, LookAt:=xlWhole, searchorder:=xlByRows) Is Nothing Then
                    room5.value = roomCell4.value
                    Exit For
                End If
            End If
        Next roomCell4
    Next room5
    
     ' add 3W schedule to diddies
    For Each room3 In Sheets("All Diddies").range("RoomsRow3WAllDiddies")
        Set newRoom = Sheets("3W Schedule").range("Rooms3WSchedule").Find(room3.value, LookAt:=xlWhole, searchorder:=xlByRows)
        ' i will start two rows below room3 in the 6:30 time slot
        i = 2
        If Not newRoom Is Nothing Then
            For j = 1 To 22
            If Not IsEmpty(newRoom.Offset(0, j)) And Not newRoom.Offset(0, j) = " " And Not newRoom Is Nothing Then
                ' put value in getNameAndProf function and return name, lunch or gray
                newString = getNameAndProf(newRoom.Offset(0, j).value)
                ' if function returns therapist name, TMG or lunch, add to diddie
                If newString = "LUNCH" Or newString <> "GRAY" Then
                    room3.Offset(i, 0).value = newString
                End If
            End If
            ' increment row on diddie schedule
            i = i + 1
        Next j
        End If
        
    Next room3
    
    
    ' add 8P schedule to diddies
    For Each room4 In Sheets("All Diddies").range("RoomsRow8PAllDiddies")
        Set newRoom = Sheets("8P Schedule").range("Rooms8PSchedule").Find(room4.value, LookAt:=xlWhole, searchorder:=xlByRows)
        ' i will start two rows below room4 in the 6:30 time slot
        i = 2
        If Not newRoom Is Nothing Then
            For j = 1 To 22
            If Not IsEmpty(newRoom.Offset(0, j)) And Not newRoom.Offset(0, j) = " " And Not newRoom Is Nothing Then
                ' put value in getNameAndProf function and return name, lunch or gray
                newString = getNameAndProf(newRoom.Offset(0, j).value)
                ' if function returns therapist name, TMG or lunch, add to diddie
                If newString = "LUNCH" Or newString <> "GRAY" Then
                    room4.Offset(i, 0).value = newString
                End If
            End If
            ' increment row on diddie schedule
            i = i + 1
        Next j
        End If
    Next room4
    
    ' add 3P schedule to diddies
    For Each room6 In Sheets("All Diddies").range("RoomsRow3PAllDiddies")
        Set newRoom = Sheets("3P Schedule").range("Rooms3PSchedule").Find(room6.value, LookAt:=xlWhole, searchorder:=xlByRows)
        ' i will start two rows below room4 in the 6:30 time slot
        i = 2
        If Not newRoom Is Nothing Then
            For j = 1 To 22
            If Not IsEmpty(newRoom.Offset(0, j)) And Not newRoom.Offset(0, j) = " " And Not newRoom Is Nothing Then
                ' put value in getNameAndProf function and return name, lunch or gray
                newString = getNameAndProf(newRoom.Offset(0, j).value)
                ' if function returns therapist name, TMG or lunch, add to diddie
                If newString = "LUNCH" Or newString <> "GRAY" Then
                    room6.Offset(i, 0).value = newString
                End If
            End If
            ' increment row on diddie schedule
            i = i + 1
        Next j
        End If
    Next room6
    
    
    ' show time the schedule was last created
    Call lastTimeCreated(timeCreated)
    ' go back to top of sheet
    ActiveWindow.ScrollRow = 1
    
    ' turn screen updating on
    Application.ScreenUpdating = True
    ' turn alerts back on
    Application.DisplayAlerts = True
    
End Sub









Attribute VB_Name = "HelperMethods"
Option Explicit

' clear contents and highlighting from All Schedules Sheet
' does not delete highlighting for Eval and Int keys above each note area
Public Sub clearAllSched()
    Sheets("All Schedules").range("AllSchedTherapistInitialsBox").value = ""
    Sheets("All Schedules").range("RoomsAllSchedules").ClearContents
    ' remove room highlighting for evals and ints
    Sheets("All Schedules").range("RoomsAllSchedules").Interior.ColorIndex = 2
    Sheets("All Schedules").range("AllSchedDataGrid").ClearContents
    ' remove time highlighting for therapist
    Sheets("All Schedules").range("AllSchedDataGrid").Interior.ColorIndex = 2
    Sheets("All Schedules").range("AllSchedNoteCells").ClearContents
    
End Sub

' clear all data  and highlighting in diddies
Public Sub clearDiddies()
Dim clearArea As range
Dim cell As range
Dim cell2 As range
Dim cell3 As range

Sheets("All Diddies").range("RoomsRow3WAllDiddies").ClearContents
Sheets("All Diddies").range("RoomsRow8PAllDiddies").ClearContents
Sheets("All Diddies").range("RoomsRow3PAllDiddies").ClearContents

' clear 3W diddies
For Each cell In Sheets("All Diddies").range("SchedCol3WAllDiddies")
    cell.MergeArea.ClearContents
    ' remove any highlighting
    cell.MergeArea.Interior.ColorIndex = 2
Next cell
' clear 8P diddies
For Each cell2 In Sheets("All Diddies").range("SchedCol8PAllDiddies")
    cell2.MergeArea.ClearContents
    ' remove any highlighting
    cell2.MergeArea.Interior.ColorIndex = 2
Next cell2
' clear 3P diddies
For Each cell3 In Sheets("All Diddies").range("SchedCol3PAllDiddies")
    cell3.MergeArea.ClearContents
    ' remove any highlighting
    cell3.MergeArea.Interior.ColorIndex = 2
Next cell3
 
End Sub

' get string value from cell and extract and return GRAY,LUNCH or name and profession
' of therapist; used to populate all diddies; will return "TMG - Procedure"
Public Function getNameAndProf(value As String) As String
    Dim initialsDict As Object
    Dim initialsCell As range
    Dim initials As Variant
    Dim nameAndProf As String
    
    ' create initials dictionary object
    Set initialsDict = CreateObject("Scripting.Dictionary")

    ' loop through all therapists to get initials, name and profession; add to dictionary
    ' key = initials; value = name + profession
    For Each initialsCell In Sheets("All Therapists").range("AllTherapistsInitials")
        If Not IsEmpty(initialsCell) And initialsCell.value <> "-" And initialsCell <> " " Then
            initials = Trim(initialsCell.value)
            nameAndProf = initialsCell.Offset(0, 3).value & " " & initialsCell.Offset(0, 2).value
            initialsDict.Add Key:=initials, Item:=nameAndProf
        End If
    Next initialsCell
    
    ' account for possible spaces and cases in lunch cell
    If Trim(LCase(value)) = "lunch" Then
        getNameAndProf = "LUNCH"
        ' reset initials dictionary
        Set initialsDict = Nothing
        Exit Function
    ElseIf isGray(value) = True Then
        getNameAndProf = "GRAY"
        ' reset initials dictionary
        Set initialsDict = Nothing
        Exit Function
    Else
        ' use returnInitials to extract initials from cell value
        If returnInitials(value) = "TMG" Then
            getNameAndProf = "TMG Procedure"
        Else
            getNameAndProf = initialsDict(returnInitials(value))
        End If
        ' reset initials dictionary
        Set initialsDict = Nothing
        Exit Function
    End If
    
    

End Function

' takes a range and displays the last time a call was made to this procedure
' used to keep track of when schedules and diddies were last created
Public Sub lastTimeCreated(range As range)
   Dim timeDate As Variant
   timeDate = Now()
   range.value = timeDate
End Sub
' takes a cell value that is not a gray option or lunch (validate before passing) and returns the initials
' THIS WILL RETURN "TMG" AS INITIALS (IT IS A PROCEDURE AND SHOULD BE LISTED
' ON DIDDIES)
Public Function returnInitials(value As String)
     ' declare variables
     Dim newValue As String
     
     ' remove trailing and initial spaces
     value = Trim(value)
     
     ' if value is > 6 characters, it is probably a note
     If Len(value) > 6 Then
        returnInitials = "NOTE"
        Exit Function
     End If

     
     If Len(value) > 3 Then ' initials in schedules are 2 or 3 characters
            If Len(value) > 5 Then
                newValue = Right(value, 3)
                newValue = Replace(newValue, " ", "")
                returnInitials = UCase(newValue)
                Exit Function
            Else ' length of value is 4 or 5
                newValue = Right(value, 2)
                returnInitials = UCase(newValue)
                Exit Function
            End If
    
        Else ' length of value is <= 3
            returnInitials = UCase(value)
        End If
End Function

' returns TRUE if value is a gray option listed in 3W, 8P or 3P
Public Function isGray(value As String) As Boolean
    Dim grayOptionsDict As Object
    Dim grayCounter As Integer
    Dim cell As range
    Dim cell2 As range
    Dim cell3 As range
    Dim newVal As String
    
    
    ' create gray options dictionary object
    Set grayOptionsDict = CreateObject("Scripting.Dictionary")
    
    ' set grayOption counter to use as arbitrary value for dictionary
    grayCounter = 1
    
    ' standardize value
    value = Replace(value, " ", "")
    value = LCase(value)
    
    ' loop through 3W gray options and add to dictionary
    For Each cell In Sheets("3W Schedule").range("GrayOptions3W")
        newVal = cell.value
        newVal = Replace(newVal, " ", "")
        newVal = LCase(newVal)
        If Not IsEmpty(cell) And cell.value <> " " And Not grayOptionsDict.Exists(newVal) Then
            grayOptionsDict.Add Key:=newVal, Item:=grayCounter
            grayCounter = grayCounter + 1
        End If
    Next cell
    
    ' loop through 8P gray options and add to dictionary
    For Each cell2 In Sheets("8P Schedule").range("GrayOptions8P")
        newVal = cell2.value
        newVal = Replace(newVal, " ", "")
        newVal = LCase(newVal)
        If Not IsEmpty(cell2) And cell2.value <> " " And Not grayOptionsDict.Exists(newVal) Then
            grayOptionsDict.Add Key:=newVal, Item:=grayCounter
            grayCounter = grayCounter + 1
        End If
    Next cell2
    
    ' loop through 3P gray options and add to dictionary
    For Each cell3 In Sheets("3P Schedule").range("GrayOptions3P")
        newVal = cell3.value
        newVal = Replace(newVal, " ", "")
        newVal = LCase(newVal)
        If Not IsEmpty(cell3) And cell3.value <> " " And Not grayOptionsDict.Exists(newVal) Then
            grayOptionsDict.Add Key:=newVal, Item:=grayCounter
            grayCounter = grayCounter + 1
        End If
    Next cell3
    
    If grayOptionsDict.Exists(value) Then
        isGray = True
    Else
        isGray = False
    End If
    
    ' reset gray options dictionary
    Set grayOptionsDict = Nothing
End Function
' copy named ranges from one workbook to another
Sub copyNames()
    Dim Source As Workbook
    Dim Target As Workbook
    Dim n As Name

    Set Source = Workbooks("SchedulingWorkbookCopy.xlsm")
    Set Target = Workbooks("Sprint1SchedulingWorkbook2.xlsm")

    For Each n In Source.Names
        Target.Names.Add Name:=n.Name, RefersTo:=n.value
    Next
End Sub
' clear contents and highlighting from Ind Schedules Sheet
' does not delete highlighting for Eval and Int key above each note area
Public Sub clearIndSched()
    Sheets("Ind Schedule").range("IndSchedRooms").ClearContents
    ' remove room highlighting for evals and ints
    Sheets("Ind Schedule").range("IndSchedRooms").Interior.ColorIndex = 2
    Sheets("Ind Schedule").range("IndSchedData").ClearContents
    ' remove time highlighting for therapist
    Sheets("Ind Schedule").range("IndSchedData").Interior.ColorIndex = 2
    Sheets("Ind Schedule").range("IndSchedNoteArea").ClearContents
    
End Sub

' apply conditional formatting to All Schedules and Ind Schedule
' takes cells for eval and int highlighting keys in All Schedules or Ind Schedule sheets, room ranges to highlight for INT and EVAL
' in All Schedules or Ind Schedule sheets, therapist's initials box range to check, and sheet name to highlight (either All Schedules or
' Ind Schedule)
Public Sub schedCondFormat(evalKey As range, intKey As range, schedRooms As range, therInitialsBox As range, schedSheet As Worksheet)
    ' declare variables
    Dim evalInt As range
    Dim roomCell As range
    Dim newEI As range
    Dim newRoom As range
    Dim evalInt8P As range
    Dim roomCell8P As range
    Dim newEI8P As range
    Dim newRoom8P As range
     Dim evalInt3P As range
    Dim roomCell3P As range
    Dim newEI3P As range
    Dim newRoom3P As range
    Dim initialsBox As range
    Dim initialsCell As range
    Dim schedCell As range
    Dim initialsValue As String
    
    ' turn off screen updating
    Application.ScreenUpdating = False
    
    ' show that EVALs are highlighted in yellow
    With evalKey
        .value = "EVAL"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 36
    End With
     ' show that INTs are highlighted in pink
    With intKey
        .value = "INT"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 38
    End With
    
    ' highlight rooms in 3W with EVALs in yellow
    For Each evalInt In Sheets("3W Schedule").range("Eval_Int_3W")
        If LCase(evalInt.value) = "eval" Then
            For Each roomCell In schedRooms
               If evalInt.Offset(0, 1).value = roomCell.value Then
                roomCell.Interior.ColorIndex = 36
               End If
            Next roomCell
        End If
    Next evalInt
        
      ' highlight rooms in 3W with INTs in pink
    For Each newEI In Sheets("3W Schedule").range("Eval_Int_3W")
        If LCase(newEI.value) = "int" Then
            For Each newRoom In schedRooms
               If newEI.Offset(0, 1).value = newRoom.value Then
                newRoom.Interior.ColorIndex = 38
               End If
            Next newRoom
        End If
    Next newEI
    
    ' highlight rooms in 8P with EVALs in yellow
    For Each evalInt8P In Sheets("8P Schedule").range("Eval_Int_8P")
        If LCase(evalInt8P.value) = "eval" Then
            For Each roomCell8P In schedRooms
               If evalInt8P.Offset(0, 1).value = roomCell8P.value Then
                roomCell8P.Interior.ColorIndex = 36
               End If
            Next roomCell8P
        End If
    Next evalInt8P
        
      ' highlight rooms in 8P with INTs in pink
    For Each newEI8P In Sheets("8P Schedule").range("Eval_Int_8P")
        If LCase(newEI8P.value) = "int" Then
            For Each newRoom8P In schedRooms
               If newEI8P.Offset(0, 1).value = newRoom8P.value Then
                newRoom8P.Interior.ColorIndex = 38
               End If
            Next newRoom8P
        End If
    Next newEI8P
    
     ' highlight rooms in 3P with EVALs in yellow
    For Each evalInt3P In Sheets("3P Schedule").range("Eval_Int_3P")
        If LCase(evalInt3P.value) = "eval" Then
            For Each roomCell3P In schedRooms
               If evalInt3P.Offset(0, 1).value = roomCell3P.value Then
                roomCell3P.Interior.ColorIndex = 36
               End If
            Next roomCell3P
        End If
    Next evalInt3P
        
      ' highlight rooms in 3P with INTs in pink
    For Each newEI3P In Sheets("3P Schedule").range("Eval_Int_3P")
        If LCase(newEI3P.value) = "int" Then
            For Each newRoom3P In schedRooms
               If newEI3P.Offset(0, 1).value = newRoom3P.value Then
                newRoom3P.Interior.ColorIndex = 38
               End If
            Next newRoom3P
        End If
    Next newEI3P
    
    ' highlight therapist's rooms on schedule in a different yellow
    For Each initialsCell In therInitialsBox
        For Each schedCell In schedSheet.range(initialsCell.Offset(3, -10), initialsCell.Offset(20, 11))
            If Not IsEmpty(schedCell) Then
                initialsValue = returnInitials(schedCell.value)
                If UCase(initialsCell.value) = initialsValue Then
                    schedCell.Interior.ColorIndex = 6
                End If
            End If
        Next schedCell
    Next initialsCell

    ' turn on screen updating
    Application.ScreenUpdating = True
End Sub
' searches for a val in sRng, if the value matches the value
'in the same row first column is returned to the cell
Public Function searchRange(val As range, sRng As range)
    Dim cel As range
    For Each cel In sRng.Cells
        If cel.value = val.value Then
            searchRange = cel.Parent.Cells(cel.Row, 1).value
            Exit Function
        End If
    Next cel
    searchRange = ""
End Function
' highlight rooms assigned to more than one OT or more than one PT
' in worksheet_change function, able to remove highlighting when duplicates
' are manually removed
Public Sub highlightDuplicateRooms(sht As Worksheet, rng As range)
    ' loop through rooms for OTs in sheet and highligh duplicates in range
    Dim cell As range
    For Each cell In rng
        'If cell.value <> "" And cell.value <> " " And Not IsEmpty(cell) Then
            If Application.WorksheetFunction.CountIf(rng, cell.value) > 1 And Not IsEmpty(cell) Then
                cell.Interior.ColorIndex = 53
                cell.Font.ColorIndex = 2
             Else
                cell.Interior.ColorIndex = 2
                cell.Font.ColorIndex = 1
            End If
       ' End If
        
        
    Next cell
End Sub

' clears contents and formatting from All Therapists rooms
Public Sub clearAllTherapistsNotesAndRooms()
    Dim allRooms As range
    Dim allNotes As range
    Dim cell As range
    Dim cell2 As range
    
    Application.ScreenUpdating = False
    
    
    Set allRooms = Sheets("All Therapists").range("AllTherapistsAllRooms")
    Set allNotes = Sheets("All Therapists").range("AllTherapistsAllNotes")
    ' clear highlighting and contents from all therapists room cells
    For Each cell In allRooms
        If cell.value <> "" And cell.value <> " " And Not IsEmpty(cell) Then
            cell.ClearContents
        End If
        If cell.Interior.ColorIndex = 53 Then
            cell.Interior.ColorIndex = 2
            cell.Font.ColorIndex = 1
        End If
        
    Next cell
    
    For Each cell2 In allNotes
        If cell2.value <> "" And cell2.value <> " " And Not IsEmpty(cell2) Then
            cell2.MergeArea.ClearContents
        End If
    Next cell2
    
     Application.ScreenUpdating = True
    
End Sub

' Used to find the name of a worksheet; accepts the name as a string.
Function GetWorksheet(shtName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = Worksheets(shtName)
End Function
' create and return dictionary of initials and address of initials in All Therapists
Public Function createInitialsDict()
 Dim initialsDict As Object
 Dim initialsCell As range
 Dim initials As String
 Dim initialsAddress As String
 
 Set initialsDict = CreateObject("Scripting.Dictionary")
 
 ' loop through all therapists to get initials and address of initials; add to dictionary
    ' key = initials; value = address
    For Each initialsCell In Sheets("All Therapists").range("AllTherapistsInitials")
        If Not IsEmpty(initialsCell) And initialsCell.value <> "-" And initialsCell <> " " Then
            initials = UCase(Trim(initialsCell.value))
            initialsAddress = initialsCell.Address
            If Not initialsDict.Exists(initials) Then
                initialsDict.Add Key:=initials, Item:=initialsAddress
            End If
        End If
    Next initialsCell
    
    Set createInitialsDict = initialsDict
    Set initialsDict = Nothing
    Exit Function
    
End Function
' create and return dictionary of rooms and wings
Public Function createRoomsDict()
    Dim roomsDict As Object
    Dim roomCell1 As range
    Dim roomCell2 As range
    Dim roomCell3 As range
    Dim roomNum As String
    Dim wing As String
    
    ' create dictionary
    Set roomsDict = CreateObject("Scripting.Dictionary")
    
    ' loop through schedules to get rooms and associate them with wing
    ' key = roomNum; value = wing
    For Each roomCell1 In Sheets("3W Schedule").range("Rooms3WSchedule")
        roomNum = roomCell1.value
        wing = "3W"
        If Not roomsDict.Exists(roomNum) Then
            roomsDict.Add Key:=roomNum, Item:=wing
        End If
    Next roomCell1
    
    For Each roomCell2 In Sheets("8P Schedule").range("Rooms8PSchedule")
        roomNum = UCase(roomCell2.value)
        wing = "8P"
        If Not roomsDict.Exists(roomNum) Then
            roomsDict.Add Key:=roomNum, Item:=wing
        End If
    Next roomCell2
    
    For Each roomCell3 In Sheets("3P Schedule").range("Rooms3PSchedule")
        roomNum = UCase(roomCell3.value)
        wing = "3P"
        If Not roomsDict.Exists(roomNum) Then
            roomsDict.Add Key:=roomNum, Item:=wing
        End If
    Next roomCell3
    
    Set createRoomsDict = roomsDict
    Set roomsDict = Nothing
    Exit Function
    
End Function

' procedure to test calls to clear and populate All Therapists
Public Sub populateAllTherapists()
    Call clearAllTherapistsNotesAndRooms
    Call getTherapistsRooms(Sheets("TherRooms3W"))
    Call getTherapistsRooms(Sheets("TherRooms8P"))
    Call getTherapistsRooms(Sheets("TherRooms3P"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
End Sub



Public Sub getLastRow(sheet As Worksheet, printCell As range)
    Dim lastRow As Long
    lastRow = sheet.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    ' print last row number in cell
    printCell.value = lastRow
End Sub

' changes value of last row cell to 0 so that addRooms will work if
' therapist form is cleared
Public Sub changeLastRow()
    Sheets("All Therapists").Unprotect Password:="Roper"
    Sheets("All Therapists").range("LastRowCell3W").value = 0
    Sheets("All Therapists").Protect Password:="Roper"
End Sub



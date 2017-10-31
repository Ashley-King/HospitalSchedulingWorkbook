Attribute VB_Name = "TherapistFormMod"
Option Explicit
' read therapist form after it is imported and put rooms into All Therapists sheet
Public Sub getTherapistsRoomsAndNotes(therapistSheet As Worksheet, Optional lastRow As Long)
    Dim targetSheet As Worksheet
    Dim therSheet As Worksheet
    Dim therShtRng As range
    Dim initialsDict As Object
    Dim roomsDict As Object
    Dim initialsCell As range
    Dim initialsStr As String
    Dim requestedRooms As String
    Dim roomsArray() As String
    Dim rm As Variant
    Dim nextCell As range
    Dim note As range
    Dim atNote As range
    Dim shtNote As String
    Dim atNoteVal As String
    Dim i As Integer
    Dim newCell As range
    Dim last As Long
  
    ' turn off screen updating
    Application.ScreenUpdating = False
    
    ' set sheets
    Set targetSheet = ActiveWorkbook.Sheets("All Therapists")
    Set therSheet = therapistSheet

    
     ' keep up with last row read on Rooms if optional parameter is passed it
    If lastRow = 0 Then
        ' read to last filled row of B
        Set therShtRng = therSheet.range("$B$2:$B" & range("$B$2").End(xlDown).Row)
    Else
        ' this means start reading forms after row lastRow of column B
        last = lastRow + 1
        ' read to last filled row of B
        Set therShtRng = therSheet.range("$B$" & last & ":$B$" & range("$B$1").End(xlDown).Row + 1)
    End If
    
    ' create  dictionary objects
    Set initialsDict = createInitialsDict()
    Set roomsDict = createRoomsDict()
    
    ' loop through rooms sheet to get initials
    ' if initials exist in dictionary (are valid), and rooms exist, get the rooms
    ' and add them to All Therapists
    For Each initialsCell In therShtRng
        ' validate initials
        If initialsDict.Exists(UCase(initialsCell.value)) Then
           ' get address of initials in All Therapists
           Set nextCell = targetSheet.range(initialsDict(UCase(initialsCell.value)))
           ' set cell for first time slot next to initials in All Therapists
           Set nextCell = nextCell.Offset(0, 4)
           ' get rooms and notes as entered
           requestedRooms = initialsCell.Offset(0, 2).value
           shtNote = initialsCell.Offset(0, 3).value
           ' remove comma and subsitute space
           requestedRooms = Replace(requestedRooms, ",", " ")
           ' substitute two spaces with one space
           requestedRooms = Replace(requestedRooms, "  ", " ")
           ' create array of rooms from string
           roomsArray() = Split(requestedRooms, " ")
           
           
        ' address of note cell in All Therapists
        Set atNote = targetSheet.range(initialsDict(UCase(initialsCell.value))).Offset(0, 25)
        ' if note is empty, add note; otherwise concatanate note
        If IsEmpty(atNote) Or atNote.value = "" Or atNote.value = "" Then
            atNote.value = shtNote
        Else
            atNoteVal = atNote.value
            ' don't duplicate notes
            If InStr(atNote, shtNote) = 0 Then
                atNote.value = atNoteVal + "; " + shtNote
            End If
        End If
           
           ' loop through rooms array; use initials in form to find address of same initials in All Therapists
           ' using initials dictionary
           For i = 0 To UBound(roomsArray)
                ' avoid skipping a cell because there's a space in the array and validate room
                If roomsArray(i) <> " " And roomsDict.Exists(UCase(roomsArray(i))) Then
                    ' for each cell in the time slots section of All Therapists beside initials
                    For Each newCell In range(nextCell, nextCell.Offset(0, 17))
                        ' don't add duplicate rooms for therapist
                        If Application.WorksheetFunction.CountIf(range(nextCell, nextCell.Offset(0, 17)), UCase(roomsArray(i))) >= 1 Then
                            Exit For
                        End If
                        If IsEmpty(newCell) Or newCell.value = "" Or newCell.value = " " Then
                            newCell.value = UCase(roomsArray(i))
                            Exit For
                        End If
                    Next newCell
                End If
           Next i
        End If

    Next initialsCell
    
    ' highlight duplicates
    Call highlightDuplicateRooms(targetSheet, targetSheet.range("AllTherapistsOTRooms"))
    Call highlightDuplicateRooms(targetSheet, targetSheet.range("AllTherapistsPTRooms"))
    
    ' reset  dictionary objects
    Set initialsDict = Nothing
    Set roomsDict = Nothing
    
    ' turn on screen updating
    Application.ScreenUpdating = True
End Sub


' Copies the Therapist Sheet from the TherapistForm.xlsm Workbook to the SchedulingWorkbook.
' If the sheet already exists, it is deleted and replaced with a new version. The copy is placed AFTER the Ind Schedule sheet.

Public Sub copyTherapistSheet(filePath As String, sheetName As String, workbookName As String)

Dim targetWorkbook As Workbook

' set filename that contains sheet to copy and targetWorkbook (which may also have
' to be changed to a string and set to a file path)

Dim Path As String
Path = filePath

Dim sName As String
sName = sheetName

Dim wName As String
wName = workbookName

Set targetWorkbook = Workbooks("SchedulingWorkbook.xlsm")

If Not GetWorksheet(sName) Is Nothing Then
    Application.DisplayAlerts = False
    Worksheets(sName).Delete
    Application.DisplayAlerts = True
End If

Workbooks.Open Filename:=Path
Sheets(sheetName).Copy After:=targetWorkbook.Sheets("Ind Schedule")
Sheets(sheetName).Name = sName
Workbooks(wName).Close savechanges:=False

ActiveWorkbook.Sheets(sName).Visible = False

End Sub

Public Sub copy3PSheet()
    
    Call copyTherapistSheet(Worksheets("File Paths").Cells(2, "B").value, "3PFormSheet", Worksheets("File Paths").Cells(2, "C").value)
    
End Sub

Public Sub copy3WSheet()

    Call copyTherapistSheet(Worksheets("File Paths").Cells(4, "B").value, "3WFormSheet", Worksheets("File Paths").Cells(4, "C").value)
    
End Sub

Public Sub copy8PSheet()

    Call copyTherapistSheet(Worksheets("File Paths").Cells(3, "B").value, "8PFormSheet", Worksheets("File Paths").Cells(3, "C").value)

End Sub


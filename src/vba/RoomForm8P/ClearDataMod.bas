Attribute VB_Name = "ClearDataMod"
Sub Clear_Output()
    ' Clear_Output Macro
    ' This clear the output information on the Rooms Sheet, and leaves it empty for the next day of use.
    Dim toPath As String
    Dim wName As String
    Dim wbTo As Workbook
    Dim wSheetName As Worksheet
    Dim wbk As Workbook
    Dim openFlag As Boolean
    
    Sheets("3WFormSheet").Select
    Range("A2:E22").Select
    ActiveWindow.SmallScroll Down:=-27
    Selection.ClearContents
    Sheets("Menu").Select
    
    
    For Each wbk In Workbooks
        If wbk.Name = "SchedulingWorkbook.xlsm" Then
            Set wbTo = Workbooks("SchedulingWorkbook.xlsm")
            wbTo.Activate
            openFlag = True
            Exit For
        Else
            Set wbTo = Workbooks.Open(Filename:="C:\Users\a\Desktop\Roper\spreadsheets\SchedulingWorkbook.xlsm")
            wbTo.Activate
            openFlag = False ' wb was not already open
            Exit For
        End If
    Next wbk

    ' reset last row read in scheduling notebook
    toPath = Workbooks("RoomForm8P.xlsm").Sheets("Menu").Range("$A$40").Value ' change name

    ' get appropriate worksheet in sheduling wb
    Set wSheetName = wbTo.Sheets("All Therapists")
    ' unprotect, update cell, protect, save and close wb if it was closed before
    wSheetName.Unprotect Password:="Roper"
    ActiveWorkbook.Sheets("All Therapists").Range("LastRowCell8P").Value = 1 ' change this
    wSheetName.Protect Password:="Roper", UserInterFaceOnly:=True
    wbTo.Save
    If openFlag = False Then
        wbTo.Close savechanges:=True ' close wb again
    End If

    End Sub

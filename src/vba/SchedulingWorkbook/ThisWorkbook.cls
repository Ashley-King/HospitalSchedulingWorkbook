VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' creates workbook open even that allows macros to run when protected but does not allow
' other modifications
Private Sub Workbook_Open()
    Dim wSheetName As Worksheet
    For Each wSheetName In Worksheets
        wSheetName.Protect Password:="Roper", UserInterFaceOnly:=True
    Next wSheetName
    ' schedule change last row cell so that addRooms will work after therapist form is cleared and repopulated
    Application.OnTime TimeValue("18:30:00"), "changeLastRow"

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As range)
    Dim cell As range
    Dim cell2 As range
    Dim otRooms As range
    Dim ptRooms As range


    ' set ranges for room cells in All Therapists (OT and PT only)
    Set otRooms = Sheets("All Therapists").range("AllTherapistsOTRooms")
    Set ptRooms = Sheets("All Therapists").range("AllTherapistsPTRooms")


    If Not (Intersect(Target, otRooms) Is Nothing) Then
        Call highlightDuplicateRooms(Sheets("All Therapists"), otRooms)
    End If
    If Not (Intersect(Target, ptRooms) Is Nothing) Then
        Call highlightDuplicateRooms(Sheets("All Therapists"), ptRooms)
    End If
    
End Sub



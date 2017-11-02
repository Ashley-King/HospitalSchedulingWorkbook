VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' change schedule when initials cell is changed and enter is pressed
Private Sub Worksheet_Change(ByVal Target As Excel.range)
    Dim vRange As range
    Dim cell As range
    
    Set vRange = Sheets("Ind Schedule").range("IndSchedInitials")
    If Union(Target, vRange).Address = vRange.Address Then
        Call createIndSched
    End If
   
End Sub


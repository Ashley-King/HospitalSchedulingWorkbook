Attribute VB_Name = "IndScheduleMod"
Option Explicit
Private rowCounter As Integer
' create schedule for individual therapist
Sub createIndSched()
    Dim cell As range
    Dim ucStr As String
    Dim cell2 As range
    Dim str1 As String
    Dim str2 As String
  
   ' clear previous data and highlighting
   Call clearIndSched
   
   ' turn off screen updating
   Application.ScreenUpdating = False

   ' look at each row in 3W, 8P & 3P schedules; if initials match initials in row
   ' copy room and schedule row and paste into all schedules


   ' initialize row counter
   rowCounter = 3
   ' initialize cell1 as therapist initials cell
   Set cell = Sheets("Ind Schedule").range("IndSchedInitials")
   ' make initials upper case if they aren't
   ucStr = UCase(cell.value)
   
    Call populateSchedules(Sheets("3W Schedule"), "SchedGrid3W", cell, rowCounter)
    Call populateSchedules(Sheets("8P Schedule"), "SchedGrid8P", cell, rowCounter)
    Call populateSchedules(Sheets("3P Schedule"), "SchedGrid3P", cell, rowCounter)
    
    ' get notes
    For Each cell2 In Sheets("All Therapists").range("AllTherapistsInitials")
        If cell2.value = UCase(range("IndSchedInitials").value) Then
            Sheets("Ind Schedule").range("IndSchedNoteRef") = cell2.Offset(0, 25).value
        End If
    Next cell2
    
    
    ' apply highlighting
   Call schedCondFormat(Sheets("Ind Schedule").range("IndSchedEvalCell"), Sheets("Ind Schedule").range("IndSchedIntCell"), _
        Sheets("Ind Schedule").range("IndSchedRooms"), Sheets("Ind Schedule").range("IndSchedInitials"), _
        Sheets("Ind Schedule"))

    ' turn off cut copy mode
    Application.CutCopyMode = False
     
   ' deselect row and go back to initials box
   Sheets("Ind Schedule").range("IndSchedInitials").Select
   
   ' turn on screen updating
   Application.ScreenUpdating = True
   
   ' deselect row and go back to initials box
   Sheets("Ind Schedule").range("IndSchedInitials").Select
   
End Sub






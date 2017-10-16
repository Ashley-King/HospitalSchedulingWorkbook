VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

    
'Path = file path where the data will be saved ******(change this to your appropriate file location)******
    Dim Path As String
    Path = "C:\temp\"
    
'ASName = The current active sheet
    Dim ASName As String
    ASName = ActiveSheet.Name
    
'FName = FileName
    Dim FName As String
    FName = Path & ASName & ".pdf"
    
'FloorNum = the range of the schedule we are currently saving (ie 3W or 8P)
'Call GetRange and Pass it the Active Sheet Name(ASName)(ie the floor schedule we are working with) and set FloorNum to the range needed
    Dim FloorNum As String
    FloorNum = GetRange(ASName)
    
'Call SaveSchedule to save a copy for use with the Schedule Posting HTML file
    Call SaveSchedule(FName, ASName, FloorNum)
    
'Call SaveArchive to save a copy of the schudle in the archive according to date
    Call SaveArchive(ASName, FloorNum)

'Enable password protection again
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
' Save the workbook
    ActiveWorkbook.Save
            
End Sub

'This Sub will save the current schedule as PDF for use with Scheduling HTML script
Private Sub SaveSchedule(FName As String, ASName As String, FloorNum As String)

    With ActiveSheet.PageSetup
        .CenterHeader = ASName & " For " & GetDate()
        .Orientation = xlLandscape
        .PrintArea = "$B$1:$X$25"
        .Zoom = False
    End With
    
   ActiveSheet.ExportAsFixedFormat _
   Type:=xlTypePDF, _
   Filename:=FName, _
   Quality:=xlQualityStandard, _
   IncludeDocProperties:=True, _
   IgnorePrintAreas:=False, _
   OpenAfterPublish:=True
   
   

End Sub
'This Sub will save the current schedule as an Archived PDF file
Private Sub SaveArchive(ASName As String, FloorNum As String)
    
    With ActiveSheet.PageSetup
        .CenterHeader = "Archived Copy: " & ASName & " For " & GetDate()
        .Orientation = xlLandscape
        .PrintArea = "$B$1:$X$25"
        .Zoom = False
    End With
    
    Dim Archive As String
    Archive = "C:\temp\Archive"
        
   ActiveSheet.ExportAsFixedFormat _
   Type:=xlTypePDF, _
   Filename:=Archive & " " & ASName & GetDate(), _
   Quality:=xlQualityStandard, _
   IncludeDocProperties:=True, _
   IgnorePrintAreas:=False, _
   OpenAfterPublish:=True

End Sub
Private Sub SaveToODS(ASName As String, FloorNum As String)


ActiveSheet.range(FloorNum).Select




End Sub
'This Sub will open the current schedule in a web-browser
Private Sub LaunchSchedulizer(ASName)

    If ASName = "8P Schedule" Then
       ThisWorkbook.FollowHyperlink ("C:\tmp\Schedulizer 8P.pdf")
    ElseIf ASName = "3W Schedule" Then
       ThisWorkbook.FollowHyperlink ("C:\tmp\Schedulizer 3W.pdf")
    End If

End Sub
'This Function will return the current date being used
Public Function GetDate() As String

    If Hour(Now()) < 17 Then
        GetDate = Format(Now(), "mm.dd.yyyy")
        Exit Function
    Else
        GetDate = Format(Now() + 1, "mm.dd.yyyy")
        Exit Function
    End If
    
End Function
'This Function will return the range needed for each of the 2 schedules
Private Function GetRange(ASName As String) As String
  If ASName = "8P Schedule" Then
        GetRange = "$B$1:$AA$32"
  End If


End Function

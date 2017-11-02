Attribute VB_Name = "SchedulingCallbacks"
Option Explicit

' These callbacks are for use in the custom ui xml file
' xml requires callback function to trigger macro

' callback for create all schedules in scheduling tab
Public Sub allSchedCallBack(control As IRibbonControl)
    Call Create_All_Schedules
End Sub

' callback for create all diddies in scheduling tab
Public Sub allDiddiesCallBack(control As IRibbonControl)
    Call Create_All_Diddies
End Sub

' callback for Import Rooms and Notes in scheduling tab
Public Sub importRooms(control As IRibbonControl)
    Call clearAllTherapistsNotesAndRooms
    Call copy3WSheet
    Call getTherapistsRoomsAndNotes(Sheets("3WFormSheet"))
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("3WFormSheet"), Sheets("All Therapists").range("LastRowCell3W"))
    Call copy8PSheet
    Call getTherapistsRoomsAndNotes(Sheets("8PFormSheet"))
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("8PFormSheet"), Sheets("All Therapists").range("LastRowCell8P"))
    Call copy3PSheet
    Call getTherapistsRoomsAndNotes(Sheets("3PFormSheet"))
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("3PFormSheet"), Sheets("All Therapists").range("LastRowCell3P"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub

' callback for Add Rooms and Notes From 3W in scheduling tab
Public Sub addRooms3W(control As IRibbonControl)
    Call copy3WSheet
    Call getTherapistsRoomsAndNotes(Sheets("3WFormSheet"), Sheets("All Therapists").range("LastRowCell3W").value)
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("3WFormSheet"), Sheets("All Therapists").range("LastRowCell3W"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub
' callback for Add Rooms and Notes From 8P in scheduling tab
Public Sub addRooms8P(control As IRibbonControl)
   Call copy8PSheet
    Call getTherapistsRoomsAndNotes(Sheets("8PFormSheet"), Sheets("All Therapists").range("LastRowCell8P").value)
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("8PFormSheet"), Sheets("All Therapists").range("LastRowCell8P"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub
' callback for Add Rooms and Notes From 3P in scheduling tab
Public Sub addRooms3P(control As IRibbonControl)
   Call copy3PSheet
    Call getTherapistsRoomsAndNotes(Sheets("3PFormSheet"), Sheets("All Therapists").range("LastRowCell3P").value)
    ' record last row read when sheet was imported
    Call getLastRow(Sheets("3PFormSheet"), Sheets("All Therapists").range("LastRowCell3P"))
    Call lastTimeCreated(Sheets("All Therapists").range("AllTherapistsTimeCreatedCell"))
    Sheets("All Therapists").Select
End Sub
' callback for Save 3W in scheduling tab
Public Sub save3W(control As IRibbonControl)
    Call Save3WSchedule
End Sub
' callback for Save 8P in scheduling tab
Public Sub save8P(control As IRibbonControl)
    Call Save8PSchedule
End Sub
' callback for Save 3P in scheduling tab
Public Sub save3P(control As IRibbonControl)
    Call Save3PSchedule
End Sub



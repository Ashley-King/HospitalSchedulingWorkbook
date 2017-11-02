Attribute VB_Name = "ExportPDFMod"
Sub Button1_Click()

Rooms.Range("A1:F18").ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\Documents\RoomsFormPDF.pdf"

End Sub
Sub Therapists_Button1_Click()

Main.Range("A1:H20").ExportAsFixedFormat xlTypePDF, Environ("UserProfile") & "\Documents\TherapistsFormPDF.pdf"


End Sub

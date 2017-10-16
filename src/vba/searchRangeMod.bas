Attribute VB_Name = "searchRangeMod"
Option Explicit

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

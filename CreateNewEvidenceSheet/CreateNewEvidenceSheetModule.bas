Attribute VB_Name = "Module1"
Option Explicit
Public Sub createNewEvidenceSheet()
    Dim newSheet As Worksheet
    Set newSheet = Worksheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Activate
    ActiveWindow.Zoom = 85
    newSheet.Cells.NumberFormat = "@"
End Sub

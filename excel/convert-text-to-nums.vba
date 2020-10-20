' From: https://www.exceltip.com/cells-ranges-rows-and-columns-in-vba/change-text-to-number-using-vba.html'
' Highlight cells that need to be update and run macro'
' Quick enough to be run on entire worksheet?  So far'

Sub ConvertTextToNumbers()
  On Error Resume Next
  Dim rSelection As Range
  Set rSelection = rSelection
  rSelection.Select

  With Selection
    Selection.NumberFormat = "General"
    .Value = .Value
  End With

  rSelection.Select
  Set rSelection = Nothing
End Sub

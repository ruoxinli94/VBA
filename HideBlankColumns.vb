Public Sub HydeEmptyColumns()

  Dim rng As Range
  Dim nLastRow As Long
  Dim nLastColumn As Integer
  Dim i As Integer
  Dim HideIt As Boolean
  Dim j As Long

  Set rng = ActiveSheet.UsedRange
  nLastRow = rng.Rows.Count + rng.row - 1
  nLastColumn = rng.Columns.Count + rng.Column - 1

  For i = 1 To nLastColumn
    HideIt = True

    For j = 2 To nLastRow
      If Not Rows(j).Hidden Then
        If Cells(j, i).Value <> "" Then
          HideIt = False
          Exit For
        End If
      End If
    Next

    Columns(i).EntireColumn.Hidden = HideIt
  Next

End Sub

' ------------------------------------------------------------------------------
' Functions
' ------------------------------------------------------------------------------
' Get number of rows in worksheet
Public Function getLastRowIndex(Worksheet As Worksheet, column As Long) As Long
  getLastRowIndex = Worksheet.Cells(Worksheet.Rows.Count, column).End(xlUp).row
End Function
' Get number of columns in worksheet
Public Function getLastColIndex(Worksheet As Worksheet, row As Long) As Long
  getLastColIndex = Worksheet.Cells(row, Worksheet.Columns.Count).End(xlToLeft).column
End Function
' Copy data
Public Sub copyData(source As Range, target As Range)
  target.Value = source.Value
End Sub
' Progress indicator
Public Sub progress(status As Integer)
  window.label.Caption  = status & "% Completed"                                ' Update label
  window.bar.Width      = status * 2                                            ' Update progress bar
  DoEvents                                                                      ' Display changes
End Sub

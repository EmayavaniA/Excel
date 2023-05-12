Option Explicit
Private Sub Worksheet_Change(ByVal Destination As Range)
Dim rngDropdown As Range
Dim oldValue As String
Dim newValue As String
Dim DelimiterType As String
DelimiterType = ", "
 
If Destination.Count > 1 Then Exit Sub
 
On Error Resume Next
Set rngDropdown = Cells.SpecialCells(xlCellTypeAllValidation)
On Error GoTo exitError
 
If rngDropdown Is Nothing Then GoTo exitError
If Not Destination.Column = 5 Then GoTo exitError

If Intersect(Destination, rngDropdown) Is Nothing Then
   'do nothing
Else
  Application.EnableEvents = False
  newValue = Destination.Value
  Application.Undo
  oldValue = Destination.Value
  Destination.Value = newValue
    If oldValue <> "" Then
    If newValue <> "" Then
        If oldValue = newValue Or _
            InStr(1, oldValue, DelimiterType & newValue) Or _
            InStr(1, oldValue, newValue & Replace(DelimiterType, " ", "")) Then
            Destination.Value = oldValue
                Else
            Destination.Value = oldValue & DelimiterType & newValue
        End If
    End If
    End If
End If
 
exitError:
  Application.EnableEvents = True
End Sub
 
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
 
End Sub

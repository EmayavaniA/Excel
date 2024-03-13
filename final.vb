Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim oldValue As String
    Dim newValue As String
    Dim DelimiterType As String
    DelimiterType = ", "
    Dim DelimiterCount As Integer
    Dim TargetType As Integer
    Dim i As Integer
    Dim arr() As String
    Dim rngDropdown As Range
    


    If Target.Count > 1 Then Exit Sub
    On Error Resume Next
    
     ' "Environment" is in column C
    Dim environmentColumn As Range
    Set environmentColumn = Range("C:C")
    
    ' "Permission set" is in column H
    Dim permissionSetColumn As Range
    Set permissionSetColumn = Range("H:H")
    
        ' "Permission set Group" is in column I
    Dim permissionSetGroupColumn As Range
    Set permissionSetGroupColumn = Range("I:I")

    Set rngDropdown = Cells.SpecialCells(xlCellTypeAllValidation)
    
    ' Check if the changed cell is in column C (Environment)
    If Not Intersect(Target, environmentColumn) Is Nothing Then
        ' Check if the old value and new value are different and if they are either UAT or PROD
        If Target.Value <> oldValue And (Target.Value = "UAT" Or Target.Value = "PROD") Then
            ' Clear Permission set and Permission set Group columns
            permissionSetColumn.Cells(Target.Row, 1).Value = ""
            permissionSetGroupColumn.Cells(Target.Row, 1).Value = ""
        End If
    End If
    
    

    If Not rngDropdown Is Nothing Then
        If Not Intersect(Target, rngDropdown) Is Nothing Then
            TargetType = Target.Validation.Type

            If TargetType = 3 Then ' is validation type is "list"
                If Not Intersect(Target, permissionSetColumn) Is Nothing Then
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    newValue = Target.Value
                    Application.Undo
                    oldValue = Target.Value
                    Target.Value = newValue
                    If oldValue <> "" Then
                        If newValue <> "" Then
                            If oldValue = newValue Or oldValue = newValue & Replace(DelimiterType, " ", "") Or oldValue = newValue & DelimiterType Then
                                oldValue = Replace(oldValue, DelimiterType, "")
                                oldValue = Replace(oldValue, Replace(DelimiterType, " ", ""), "")
                                Target.Value = oldValue
                            ElseIf InStr(1, oldValue, DelimiterType & newValue) Then
                                arr = Split(oldValue, DelimiterType)
                                If Not IsError(Application.Match(newValue, arr, 0)) = 0 Then
                                    Target.Value = oldValue & DelimiterType & newValue
                                Else
                                    Target.Value = ""
                                    For i = 0 To UBound(arr)
                                        If arr(i) <> newValue Then
                                            Target.Value = Target.Value & arr(i) & DelimiterType
                                        End If
                                    Next i
                                    Target.Value = Left(Target.Value, Len(Target.Value) - Len(DelimiterType))
                                End If
                            ElseIf InStr(1, oldValue, newValue & Replace(DelimiterType, " ", "")) Then
                                oldValue = Replace(oldValue, newValue, "")
                                Target.Value = oldValue
                            Else
                                Target.Value = oldValue & DelimiterType & newValue
                            End If
                            Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", "") & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", ""))
                            Target.Value = Replace(Target.Value, DelimiterType & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", ""))
                            If Target.Value <> "" Then
                                If Right(Target.Value, 2) = DelimiterType Then
                                    Target.Value = Left(Target.Value, Len(Target.Value) - 2)
                                End If
                            End If
                            If InStr(1, Target.Value, DelimiterType) = 1 Then
                                Target.Value = Replace(Target.Value, DelimiterType, "", 1, 1)
                            End If
                            If InStr(1, Target.Value, Replace(DelimiterType, " ", "")) = 1 Then
                                Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", ""), "", 1, 1)
                            End If
                            DelimiterCount = 0
                            For i = 1 To Len(Target.Value)
                                If InStr(i, Target.Value, Replace(DelimiterType, " ", "")) Then
                                    DelimiterCount = DelimiterCount + 1
                                End If
                            Next i
                            If DelimiterCount = 1 Then
                                Target.Value = Replace(Target.Value, DelimiterType, "")
                                Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", ""), "")
                            End If
                        End If
                    End If
                    Application.EnableEvents = True
                    Application.ScreenUpdating = True
                End If
            End If
        End If
    End If
    
    If Not rngDropdown Is Nothing Then
        If Not Intersect(Target, rngDropdown) Is Nothing Then
            TargetType = Target.Validation.Type

            If TargetType = 3 Then ' is validation type is "list"
                If Not Intersect(Target, permissionSetGroupColumn) Is Nothing Then
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False
                    newValue = Target.Value
                    Application.Undo
                    oldValue = Target.Value
                    Target.Value = newValue
                    If oldValue <> "" Then
                        If newValue <> "" Then
                            If oldValue = newValue Or oldValue = newValue & Replace(DelimiterType, " ", "") Or oldValue = newValue & DelimiterType Then
                                oldValue = Replace(oldValue, DelimiterType, "")
                                oldValue = Replace(oldValue, Replace(DelimiterType, " ", ""), "")
                                Target.Value = oldValue
                            ElseIf InStr(1, oldValue, DelimiterType & newValue) Then
                                arr = Split(oldValue, DelimiterType)
                                If Not IsError(Application.Match(newValue, arr, 0)) = 0 Then
                                    Target.Value = oldValue & DelimiterType & newValue
                                Else
                                    Target.Value = ""
                                    For i = 0 To UBound(arr)
                                        If arr(i) <> newValue Then
                                            Target.Value = Target.Value & arr(i) & DelimiterType
                                        End If
                                    Next i
                                    Target.Value = Left(Target.Value, Len(Target.Value) - Len(DelimiterType))
                                End If
                            ElseIf InStr(1, oldValue, newValue & Replace(DelimiterType, " ", "")) Then
                                oldValue = Replace(oldValue, newValue, "")
                                Target.Value = oldValue
                            Else
                                Target.Value = oldValue & DelimiterType & newValue
                            End If
                            Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", "") & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", ""))
                            Target.Value = Replace(Target.Value, DelimiterType & Replace(DelimiterType, " ", ""), Replace(DelimiterType, " ", ""))
                            If Target.Value <> "" Then
                                If Right(Target.Value, 2) = DelimiterType Then
                                    Target.Value = Left(Target.Value, Len(Target.Value) - 2)
                                End If
                            End If
                            If InStr(1, Target.Value, DelimiterType) = 1 Then
                                Target.Value = Replace(Target.Value, DelimiterType, "", 1, 1)
                            End If
                            If InStr(1, Target.Value, Replace(DelimiterType, " ", "")) = 1 Then
                                Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", ""), "", 1, 1)
                            End If
                            DelimiterCount = 0
                            For i = 1 To Len(Target.Value)
                                If InStr(i, Target.Value, Replace(DelimiterType, " ", "")) Then
                                    DelimiterCount = DelimiterCount + 1
                                End If
                            Next i
                            If DelimiterCount = 1 Then
                                Target.Value = Replace(Target.Value, DelimiterType, "")
                                Target.Value = Replace(Target.Value, Replace(DelimiterType, " ", ""), "")
                            End If
                        End If
                    End If
                    Application.EnableEvents = True
                    Application.ScreenUpdating = True
                End If
            End If
        End If
    End If
    Dim rng As Range
    Dim cell As Range
    Set rng = Intersect(Target, Me.Range("F:J")) ' Monitor changes in columns F to J
    If Not rng Is Nothing Then
        For Each cell In rng
            If cell.Column = 10 Then ' Action column
                Dim rowNumber As Long
                rowNumber = cell.Row
                
                If cell.Value = "ADD" Or cell.Value = "DELETE" Then
                    If Me.Cells(rowNumber, 7).Value <> "" Then
                        If MsgBox("For ADD/DELETE Action, PS or PS Group or combination of both coulmns are Mandatory .Hence clearing Reference Email", vbYesNo + vbQuestion, "Alert") = vbYes Then
                            ' Clear and disable column G
                            Me.Cells(rowNumber, 7).ClearContents
                            Me.Cells(rowNumber, 7).Locked = True
                            ' Enable columns I and H
                            Me.Cells(rowNumber, 8).Locked = False
                            Me.Cells(rowNumber, 9).Locked = False
                        Else
                            ' Revert the change
                            Application.EnableEvents = False
                            cell.Value = ""
                            Application.EnableEvents = True
                        End If
                    End If
                
                ElseIf cell.Value = "MIMIC" Then
                    If Me.Cells(rowNumber, 8).Value <> "" Or Me.Cells(rowNumber, 9).Value <> "" Then
                        If MsgBox("For MIMIC Action, Reference Email coulmn is Mandatory .Hence clearing PS and PSGroup", vbYesNo + vbQuestion, "Alert") = vbYes Then
                            ' Clear and disable columns  I and H
                            Me.Cells(rowNumber, 9).ClearContents
                            Me.Cells(rowNumber, 9).Locked = True
                            Me.Cells(rowNumber, 8).ClearContents
                            Me.Cells(rowNumber, 8).Locked = True
                        Else
                            ' Revert the change
                            Application.EnableEvents = False
                            cell.Value = ""
                            Application.EnableEvents = True
                        End If
                    End If
                End If
                
                ' Lock columns I, G, H, and the Action column after initial change
                If cell.Value <> "ADD" Or cell.Value <> "DELETE" Then
                    Me.Cells(rowNumber, 7).Locked = True
                ElseIf cell.Value <> "MIMIC" Then
                    Me.Cells(rowNumber, 8).Locked = True
                    Me.Cells(rowNumber, 9).Locked = True
                End If
                Me.Cells(rowNumber, 7).Locked = True
                Me.Cells(rowNumber, 8).Locked = True
                Me.Cells(rowNumber, 9).Locked = True
                Exit For
            ElseIf cell.Column = 7 And Me.Cells(cell.Row, 10).Value = "ADD" Or Me.Cells(cell.Row, 10).Value = "DELETE" Then
                ' If column G is edited and Action is still ADD, clear and disable it
                Application.EnableEvents = False
                Me.Cells(cell.Row, 7).ClearContents
                Me.Cells(cell.Row, 7).Locked = True
                Application.EnableEvents = True
            ElseIf cell.Column = 8 And Me.Cells(cell.Row, 10).Value = "MIMIC" Then
                ' If column H is edited and Action is still ADD, clear and disable it
                Application.EnableEvents = False
                Me.Cells(cell.Row, 8).ClearContents
                Me.Cells(cell.Row, 8).Locked = True
                Application.EnableEvents = True
                Exit For
            ElseIf cell.Column = 9 And Me.Cells(cell.Row, 10).Value = "MIMIC" Then
                ' If column I is edited and Action is still ADD, clear and disable it
                Application.EnableEvents = False
                Me.Cells(cell.Row, 9).ClearContents
                Me.Cells(cell.Row, 9).Locked = True
                Application.EnableEvents = True
                Exit For
            End If
        Next cell
    End If
End Sub


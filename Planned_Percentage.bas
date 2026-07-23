Option Explicit

Public Sub Planner_PlannedValue(control As IRibbonControl)
    AddPlannedValueColumn
End Sub

Private Sub AddPlannedValueColumn()
    Dim ws As Worksheet
    Dim insertCol As Long, resultCol As Long
    Dim weightCol As Long, planPctCol As Long
    Dim lastRow As Long
    Dim r As Long, k As Long
    Dim refs As String
    Dim level As Long
    Dim nxtWBS As Long, nxtLvl0 As Long, nxtLvl1 As Long, nxtLvl2 As Long, nxtLvl3 As Long
    Dim resultL As String
    Dim hasChildren As Boolean

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    insertCol = AskColumnByText(ws, _
        "Where do you want to insert the Planned Value column?" & vbCrLf & _
        "NB: Use the Coloumn name like H,L or K", "G")
    If insertCol = 0 Then Exit Sub

    ws.Columns(insertCol).Insert Shift:=xlToRight
    resultCol = insertCol
    ws.Cells(1, resultCol).Value = "Planned Value"

    weightCol = AskColumnByText(ws, _
        "Which column is Weightage?" & vbCrLf & _
        "NB: Specify the Coloumn name like H,L or K", "H")
    If weightCol = 0 Then Exit Sub

    planPctCol = AskColumnByText(ws, _
        "Which column is Planned Percentage?" & vbCrLf & _
        "NB: Use the Coloumn name like H,L or K", "I")
    If planPctCol = 0 Then Exit Sub

    resultL = ColLetter(resultCol)

    Application.ScreenUpdating = False

    ' Activity rows = Weightage * Planned Percentage
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
ws.Cells(r, resultCol).Formula = "=IFERROR(" & _
    ws.Cells(r, weightCol).Address(False, True) & "*" & _
    ws.Cells(r, planPctCol).Address(False, False) & ",0)"
        Else
            ws.Cells(r, resultCol).ClearContents
        End If
    Next r

    ' Pass 1: Level 4 (red) = sum detail rows until next WBS row
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 4 Then
                nxtWBS = NextWBSRow(ws, r, lastRow)
                If nxtWBS > r + 1 Then
                    ws.Cells(r, resultCol).Formula = "=SUM(" & resultL & (r + 1) & ":" & resultL & (nxtWBS - 1) & ")"
                Else
                    ws.Cells(r, resultCol).Value = 0
                End If
            End If
        End If
    Next r

    ' Pass 2: Level 3 (light orange) = sum level 4 rows if they exist, otherwise sum detail rows
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 3 Then
                nxtLvl3 = NextLevelRow(ws, r, lastRow, 3)
                refs = ""
                hasChildren = False

                For k = r + 1 To nxtLvl3 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 4 Then
                            hasChildren = True
                            If refs = "" Then
                                refs = resultL & k
                            Else
                                refs = refs & "," & resultL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, resultCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, resultCol).Formula = "=SUM(" & resultL & (r + 1) & ":" & resultL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, resultCol).Value = 0
                    End If
                End If
            End If
        End If
    Next r

    ' Pass 3: Level 2 (light green) = sum level 3 rows if they exist, otherwise sum detail rows
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 2 Then
                nxtLvl2 = NextLevelRow(ws, r, lastRow, 2)
                refs = ""
                hasChildren = False

                For k = r + 1 To nxtLvl2 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 3 Then
                            hasChildren = True
                            If refs = "" Then
                                refs = resultL & k
                            Else
                                refs = refs & "," & resultL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, resultCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, resultCol).Formula = "=SUM(" & resultL & (r + 1) & ":" & resultL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, resultCol).Value = 0
                    End If
                End If
            End If
        End If
    Next r

    ' Pass 4: Level 1 (light blue) = sum level 2 rows until next level 1
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            If GetIndentLevel(CStr(ws.Cells(r, 1).Value)) = 1 Then
                nxtLvl1 = NextLevelRow(ws, r, lastRow, 1)
                refs = ""

                For k = r + 1 To nxtLvl1 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 2 Then
                            If refs = "" Then
                                refs = resultL & k
                            Else
                                refs = refs & "," & resultL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, resultCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, resultCol).Value = 0
                End If
            End If
        End If
    Next r

    ' Pass 5: Level 0 (blue) = sum level 1 rows until next level 0
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            If GetIndentLevel(CStr(ws.Cells(r, 1).Value)) = 0 Then
                nxtLvl0 = NextLevelRow(ws, r, lastRow, 0)
                refs = ""

                For k = r + 1 To nxtLvl0 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 1 Then
                            If refs = "" Then
                                refs = resultL & k
                            Else
                                refs = refs & "," & resultL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, resultCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, resultCol).Value = 0
                End If
            End If
        End If
    Next r

    ws.Columns(resultCol).NumberFormat = "0.00%"
    Application.ScreenUpdating = True

    MsgBox "Planned Value completed.", vbInformation, "Planner"
End Sub

Private Function AskColumnByText(ByVal ws As Worksheet, ByVal promptText As String, ByVal defaultValue As String) As Long
    Dim s As String, col As Long

    s = Trim$(InputBox(promptText, "Planner", defaultValue))
    If s = "" Then Exit Function

    If IsNumeric(s) Then
        col = CLng(s)
        If col >= 1 And col <= ws.Columns.Count Then
            AskColumnByText = col
            Exit Function
        End If
    End If

    On Error Resume Next
    col = ws.Range(UCase$(s) & "1").Column
    On Error GoTo 0
    If col > 0 Then
        AskColumnByText = col
        Exit Function
    End If

    AskColumnByText = FindHeaderColumn(ws, s)
    If AskColumnByText = 0 Then MsgBox "Column not found.", vbExclamation, "Planner"
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), Trim$(headerText), vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function IsWBSRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    IsWBSRow = (Trim$(CStr(ws.Cells(r, 1).Value)) <> "" And Trim$(CStr(ws.Cells(r, 2).Value)) = "")
End Function

Private Function NextWBSRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastRow As Long) As Long
    Dim i As Long
    For i = r + 1 To lastRow
        If IsWBSRow(ws, i) Then
            NextWBSRow = i
            Exit Function
        End If
    Next i
    NextWBSRow = lastRow + 1
End Function

Private Function NextLevelRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastRow As Long, ByVal targetLevel As Long) As Long
    Dim i As Long
    For i = r + 1 To lastRow
        If IsWBSRow(ws, i) Then
            If GetIndentLevel(CStr(ws.Cells(i, 1).Value)) = targetLevel Then
                NextLevelRow = i
                Exit Function
            End If
        End If
    Next i
    NextLevelRow = lastRow + 1
End Function

Private Function GetIndentLevel(ByVal txt As String) As Long
    Dim leadingSpaces As Long
    leadingSpaces = Len(txt) - Len(LTrim$(txt))
    GetIndentLevel = Int(leadingSpaces / 2)
End Function

Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(False, False), "1")(0)
End Function

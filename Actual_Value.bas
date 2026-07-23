Option Explicit

Public Sub Planner_Weightage(control As IRibbonControl)
    AddWeightageColumnWithBudgetRollup
End Sub

Private Sub AddWeightageColumnWithBudgetRollup()
    Dim ws As Worksheet
    Dim insertCol As Long, weightCol As Long, budgetCol As Long
    Dim lastRow As Long, r As Long, k As Long
    Dim totalCell As Range
    Dim refs As String, budgetL As String, weightL As String
    Dim level As Long, nxtWBS As Long, nxtLvl0 As Long, nxtLvl1 As Long, nxtLvl2 As Long, nxtLvl3 As Long
    Dim hasChildren As Boolean

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    insertCol = AskColumnByText(ws, _
        "Where do you want to insert the Weightage column?" & vbCrLf & _
        "NB: Use the Column name like H,L or K", "G")
    If insertCol = 0 Then Exit Sub

    ws.Columns(insertCol).Insert Shift:=xlToRight
    weightCol = insertCol
    ws.Cells(1, weightCol).Value = "Weightage"

    budgetCol = AskColumnByText(ws, _
        "Which column is Budgeted value?" & vbCrLf & _
        "NB: Use the Column name like H,L or K", "F")
    If budgetCol = 0 Then Exit Sub

    budgetL = ColLetter(budgetCol)
    weightL = ColLetter(weightCol)

    On Error Resume Next
    Set totalCell = Application.InputBox( _
        Prompt:="Select TOTAL Budgeted cell (denominator), e.g. F2", _
        Title:="Weightage", Type:=8)
    On Error GoTo 0

    If totalCell Is Nothing Then Exit Sub
    If Not totalCell.Worksheet Is ws Then
        MsgBox "Please select total cell from active sheet.", vbExclamation, "Planner"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ===============================
    ' BUDGET ROLL-UP
    ' ===============================

    ' Pass 1: Level 4 (red) = sum detail rows until next WBS row
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 4 Then
                nxtWBS = NextWBSRow(ws, r, lastRow)
                If nxtWBS > r + 1 Then
                    ws.Cells(r, budgetCol).Formula = "=SUM(" & budgetL & (r + 1) & ":" & budgetL & (nxtWBS - 1) & ")"
                Else
                    ws.Cells(r, budgetCol).Value = 0
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
                                refs = budgetL & k
                            Else
                                refs = refs & "," & budgetL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, budgetCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, budgetCol).Formula = "=SUM(" & budgetL & (r + 1) & ":" & budgetL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, budgetCol).Value = 0
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
                                refs = budgetL & k
                            Else
                                refs = refs & "," & budgetL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, budgetCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, budgetCol).Formula = "=SUM(" & budgetL & (r + 1) & ":" & budgetL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, budgetCol).Value = 0
                    End If
                End If
            End If
        End If
    Next r

    ' Pass 4: Level 1 (light blue) = sum level 2 rows until next level 1
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 1 Then
                nxtLvl1 = NextLevelRow(ws, r, lastRow, 1)
                refs = ""

                For k = r + 1 To nxtLvl1 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 2 Then
                            If refs = "" Then
                                refs = budgetL & k
                            Else
                                refs = refs & "," & budgetL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, budgetCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, budgetCol).Value = 0
                End If
            End If
        End If
    Next r

    ' Pass 5: Level 0 (blue) = sum level 1 rows until next level 0
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 0 Then
                nxtLvl0 = NextLevelRow(ws, r, lastRow, 0)
                refs = ""

                For k = r + 1 To nxtLvl0 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 1 Then
                            If refs = "" Then
                                refs = budgetL & k
                            Else
                                refs = refs & "," & budgetL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, budgetCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, budgetCol).Value = 0
                End If
            End If
        End If
    Next r

    ' ===============================
    ' WEIGHTAGE FOR ACTIVITY ROWS
    ' ===============================
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
            ws.Cells(r, weightCol).Formula = "=IFERROR(" & _
                ws.Cells(r, budgetCol).Address(False, False) & "/" & totalCell.Address(True, True) & ",0)"
        Else
            ws.Cells(r, weightCol).ClearContents
        End If
    Next r

    ' ===============================
    ' WEIGHTAGE ROLL-UP
    ' ===============================

    ' Pass 1: Level 4 (red) = sum activity weightages
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 4 Then
                nxtWBS = NextWBSRow(ws, r, lastRow)
                If nxtWBS > r + 1 Then
                    ws.Cells(r, weightCol).Formula = "=SUM(" & weightL & (r + 1) & ":" & weightL & (nxtWBS - 1) & ")"
                Else
                    ws.Cells(r, weightCol).Value = 0
                End If
            End If
        End If
    Next r

    ' Pass 2: Level 3 (light orange) = sum level 4 rows if they exist, otherwise sum activity weightages
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
                                refs = weightL & k
                            Else
                                refs = refs & "," & weightL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, weightCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, weightCol).Formula = "=SUM(" & weightL & (r + 1) & ":" & weightL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, weightCol).Value = 0
                    End If
                End If
            End If
        End If
    Next r

    ' Pass 3: Level 2 (light green) = sum level 3 rows if they exist, otherwise sum activity weightages
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
                                refs = weightL & k
                            Else
                                refs = refs & "," & weightL & k
                            End If
                        End If
                    End If
                Next k

                If hasChildren Then
                    ws.Cells(r, weightCol).Formula = "=SUM(" & refs & ")"
                Else
                    nxtWBS = NextWBSRow(ws, r, lastRow)
                    If nxtWBS > r + 1 Then
                        ws.Cells(r, weightCol).Formula = "=SUM(" & weightL & (r + 1) & ":" & weightL & (nxtWBS - 1) & ")"
                    Else
                        ws.Cells(r, weightCol).Value = 0
                    End If
                End If
            End If
        End If
    Next r

    ' Pass 4: Level 1 (light blue) = sum level 2 weightages
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 1 Then
                nxtLvl1 = NextLevelRow(ws, r, lastRow, 1)
                refs = ""

                For k = r + 1 To nxtLvl1 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 2 Then
                            If refs = "" Then
                                refs = weightL & k
                            Else
                                refs = refs & "," & weightL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, weightCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, weightCol).Value = 0
                End If
            End If
        End If
    Next r

    ' Pass 5: Level 0 (blue) = sum level 1 weightages
    For r = 2 To lastRow
        If IsWBSRow(ws, r) Then
            level = GetIndentLevel(CStr(ws.Cells(r, 1).Value))
            If level = 0 Then
                nxtLvl0 = NextLevelRow(ws, r, lastRow, 0)
                refs = ""

                For k = r + 1 To nxtLvl0 - 1
                    If IsWBSRow(ws, k) Then
                        If GetIndentLevel(CStr(ws.Cells(k, 1).Value)) = 1 Then
                            If refs = "" Then
                                refs = weightL & k
                            Else
                                refs = refs & "," & weightL & k
                            End If
                        End If
                    End If
                Next k

                If refs <> "" Then
                    ws.Cells(r, weightCol).Formula = "=SUM(" & refs & ")"
                Else
                    ws.Cells(r, weightCol).Value = 0
                End If
            End If
        End If
    Next r

    ws.Columns(weightCol).NumberFormat = "0.00%"
    Application.ScreenUpdating = True

    MsgBox "Weightage Percentaged completed Succesfully .", vbInformation, "Planner"
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

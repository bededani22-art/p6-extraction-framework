Option Explicit

Public Sub Planner_ActualPercentage(control As IRibbonControl)
    If Not ShowActualPercentageHelp() Then Exit Sub
    AddActualPercentageColumn
End Sub

Private Function ShowActualPercentageHelp() As Boolean
    Dim pref As String

    pref = GetSetting("PlannerAddin", "Hints", "ActualPercentage", "Show")

    If pref <> "Hide" Then
        frmActualPercentageHelp.Show

        If frmActualPercentageHelp.UserChoice <> "CONTINUE" Then
            Unload frmActualPercentageHelp
            ShowActualPercentageHelp = False
            Exit Function
        End If

        Unload frmActualPercentageHelp
    End If

    ShowActualPercentageHelp = True
End Function

Private Sub AddActualPercentageColumn()
    Dim ws As Worksheet
    Dim insertCol As Long, resultCol As Long
    Dim totalQtyCol As Long, actualQtyCol As Long
    Dim lastRow As Long
    Dim r As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    insertCol = AskColumnByText(ws, _
        "Where do you want to insert the Actual Percentage column?" & vbCrLf & _
        "NB: Use the Coloumn name like H,L or K", "G")
    If insertCol = 0 Then Exit Sub

    ws.Columns(insertCol).Insert Shift:=xlToRight
    resultCol = insertCol
    ws.Cells(1, resultCol).Value = "Actual Percentage"

    totalQtyCol = AskColumnByText(ws, _
        "Which column is Total Quantity?" & vbCrLf & _
        "NB: Specify the Coloumn name like H,L or K", "F")
    If totalQtyCol = 0 Then Exit Sub

    actualQtyCol = AskColumnByText(ws, _
        "Which column is Actual Quantity?" & vbCrLf & _
        "NB: Specify the Coloumn name like H,L or K", "G")
    If actualQtyCol = 0 Then Exit Sub

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
            ws.Cells(r, resultCol).Formula = "=IFERROR(" & _
                ws.Cells(r, actualQtyCol).Address(False, False) & "/" & _
                ws.Cells(r, totalQtyCol).Address(False, True) & ",0)"
        Else
            ws.Cells(r, resultCol).ClearContents
        End If
    Next r

    ws.Columns(resultCol).NumberFormat = "0.00%"
    Application.ScreenUpdating = True

    MsgBox "Actual Percentage completed.", vbInformation, "Planner"
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

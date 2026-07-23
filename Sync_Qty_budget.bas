Option Explicit

Public Sub Planner_PlannedPercentage(control As IRibbonControl)
    If Not ShowPlannedPercentageHelp() Then Exit Sub
    AddPlannedPercentageColumn
End Sub

Private Function ShowPlannedPercentageHelp() As Boolean
    Dim pref As String

    pref = GetSetting("PlannerAddin", "Hints", "PlannedPercentage", "Show")

    If pref <> "Hide" Then
        frmPlannedPercentageHelp.Show

        If frmPlannedPercentageHelp.UserChoice <> "CONTINUE" Then
            Unload frmPlannedPercentageHelp
            ShowPlannedPercentageHelp = False
            Exit Function
        End If

        Unload frmPlannedPercentageHelp
    End If

    ShowPlannedPercentageHelp = True
End Function

Private Sub AddPlannedPercentageColumn()
    Dim ws As Worksheet
    Dim insertCol As Long, planCol As Long
    Dim startCol As Long, finishCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim dateCell As Range
    Dim f As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    insertCol = AskColumnByText(ws, _
        "Where do you want to insert the Planned Percentage column?" & vbCrLf & _
        "NB: Use the Coloumn name like H,L or K", "G")
    If insertCol = 0 Then Exit Sub

    ws.Columns(insertCol).Insert Shift:=xlToRight
    planCol = insertCol
    ws.Cells(1, planCol).Value = "Planned Percentage"

    startCol = AskColumnByText(ws, _
        "Which column is the Start date?" & vbCrLf & _
        "NB: Specify the Coloumn name like H,L or K", "D")
    If startCol = 0 Then Exit Sub

    finishCol = AskColumnByText(ws, _
        "Which column is the Finsh date?" & vbCrLf & _
        "NB: Use the Coloumn name like H,L or K", "E")
    If finishCol = 0 Then Exit Sub

    On Error Resume Next
    Set dateCell = Application.InputBox( _
        Prompt:="Select the specific date cell that you want to calculate the Planned Value ?                                          ,N.B :Use date Data like 4-Apr-2030 ", _
        Title:="Planned Percentage", Type:=8)
    On Error GoTo 0

    If dateCell Is Nothing Then Exit Sub
    If Not dateCell.Worksheet Is ws Then
        MsgBox "Please select the date cell from the active sheet.", vbExclamation, "Planner"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
            f = "=IFERROR(IF(AND(" & ws.Cells(r, startCol).Address(True, True) & "<=" & dateCell.Address(False, False) & "," & _
                dateCell.Address(False, False) & "<=" & ws.Cells(r, finishCol).Address(True, True) & "),(" & _
                dateCell.Address(False, False) & "-" & ws.Cells(r, startCol).Address(True, True) & "+1)/(" & _
                ws.Cells(r, finishCol).Address(True, True) & "-" & ws.Cells(r, startCol).Address(True, True) & "+1)," & _
                "IF(" & ws.Cells(r, finishCol).Address(True, True) & "<=" & dateCell.Address(False, False) & ",1," & _
                "IF(" & ws.Cells(r, startCol).Address(True, True) & "<=" & dateCell.Address(False, False) & ",(" & _
                dateCell.Address(False, False) & "-" & ws.Cells(r, startCol).Address(True, True) & "+1)/(" & _
                ws.Cells(r, finishCol).Address(True, True) & "-" & ws.Cells(r, startCol).Address(True, True) & "+1),0))),0)"
            ws.Cells(r, planCol).Formula = f
        Else
            ws.Cells(r, planCol).ClearContents
        End If
    Next r

    ws.Columns(planCol).NumberFormat = "0.00%"
    Application.ScreenUpdating = True

    MsgBox "Planned Percentage for the activity has been completed successfully.", vbInformation, "Planner"
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

Private Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(Cells(1, colNum).Address(False, False), "1")(0)
End Function

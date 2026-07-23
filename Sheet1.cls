Option Explicit

Public Sub Planner_ARemover(control As IRibbonControl)
    RemoveTrailingAFromStartFinish
End Sub

Public Sub RemoveTrailingAFromStartFinish()
    Dim ws As Worksheet
    Dim startCol As Long, finishCol As Long
    Dim lastRow As Long, r As Long
    Dim changed As Long
    
    Set ws = ActiveSheet
    
    startCol = AskUserForColumn(ws, "Select the coloumn where there is Start Date ?")
    If startCol = 0 Then Exit Sub
    
    finishCol = AskUserForColumn(ws, "Select the coloumn where there is Finish Date ?")
    If finishCol = 0 Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, finishCol).End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.Count, finishCol).End(xlUp).Row
    End If
    
    For r = 2 To lastRow
        changed = changed + CleanOneCell(ws.Cells(r, startCol))
        changed = changed + CleanOneCell(ws.Cells(r, finishCol))
    Next r
    
    MsgBox changed & " date cell(s) cleaned.", vbInformation, "Planner"
End Sub

Private Function AskUserForColumn(ByVal ws As Worksheet, ByVal promptText As String) As Long
    Dim picked As Range
    
    On Error Resume Next
    Set picked = Application.InputBox( _
        Prompt:=promptText, _
        Title:="Planner - A remover", _
        Type:=8)
    On Error GoTo 0
    
    If picked Is Nothing Then
        MsgBox "Operation cancelled.", vbExclamation, "Planner"
        Exit Function
    End If
    
    If Not picked.Worksheet Is ws Then
        MsgBox "Please select a cell from the active sheet only.", vbExclamation, "Planner"
        Exit Function
    End If
    
    AskUserForColumn = picked.Column
End Function

Private Function CleanOneCell(ByVal c As Range) As Long
    Dim v As String, core As String
    
    v = Trim$(CStr(c.Value))
    If Len(v) < 3 Then Exit Function
    
    ' Remove trailing " A" only when remaining text is a valid date
    If UCase$(Right$(v, 2)) = " A" Then
        core = Trim$(Left$(v, Len(v) - 2))
        If IsDate(core) Then
            c.Value = CDate(core)
            c.NumberFormat = "dd-mmm-yy"
            CleanOneCell = 1
        End If
    End If
End Function

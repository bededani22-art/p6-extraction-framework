Option Explicit

Public Sub Planner_WBSColoring(control As IRibbonControl)
    ApplyP6WBSColoring
End Sub

Public Sub ApplyP6WBSColoring()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, level As Long
    Dim idText As String, actName As String
    Dim rowRange As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False

    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        .Interior.Pattern = xlNone
        .Font.Bold = False
        .Font.Color = RGB(0, 0, 0)
    End With

    For r = 2 To lastRow
        idText = CStr(ws.Cells(r, 1).Value)
        actName = Trim$(CStr(ws.Cells(r, 2).Value))
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))

        If Trim$(idText) <> "" And actName = "" Then
            level = GetIndentLevel(idText)
            rowRange.Font.Bold = True

            Select Case level
                Case 0
                    rowRange.Interior.Color = RGB(91, 155, 213)
                    rowRange.Font.Color = RGB(255, 255, 255)
                Case 1
                    rowRange.Interior.Color = RGB(221, 235, 247)
                Case 2
                    rowRange.Interior.Color = RGB(226, 239, 218)
                Case 3
                    rowRange.Interior.Color = RGB(255, 242, 204)
                Case 4
                    rowRange.Interior.Color = RGB(244, 204, 204)
                Case Else
                    rowRange.Interior.Color = RGB(234, 234, 234)
            End Select
        End If
    Next r

    Application.ScreenUpdating = True
    MsgBox "P6 WBS coloring applied successfully.", vbInformation, "Planner"
End Sub

Private Function GetIndentLevel(ByVal txt As String) As Long
    Dim leadingSpaces As Long
    leadingSpaces = Len(txt) - Len(LTrim$(txt))
    GetIndentLevel = Int(leadingSpaces / 2)
End Function

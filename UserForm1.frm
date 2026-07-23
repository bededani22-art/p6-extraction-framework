Option Explicit

Private Const PLANNER_LOG_SHEET As String = "Planner_Log"

' =========================
' Ribbon Callbacks
' =========================
Public Sub Planner_SyncByActivityID(control As IRibbonControl)
    SyncByActivityID
End Sub

Public Sub Planner_ShowBlankActivities(control As IRibbonControl)
    HighlightLoggedRows "Blank", RGB(255, 255, 0)
End Sub

Public Sub Planner_ShowMissingActivities(control As IRibbonControl)
    HighlightLoggedRows "Missing", RGB(255, 199, 206)
End Sub

Public Sub Planner_ShowDuplicateActivities(control As IRibbonControl)
    HighlightLoggedRows "Duplicate", RGB(255, 235, 156)
End Sub

Public Sub Planner_ShowUpdatedRows(control As IRibbonControl)
    HighlightLoggedRows "Updated", RGB(198, 239, 206)
End Sub

Public Sub Planner_ClearReviewColors(control As IRibbonControl)
    ClearReviewColors
End Sub

' =========================
' Main Sync
' =========================
Private Sub SyncByActivityID()
    Dim destWb As Workbook
    Dim destWs As Worksheet
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim srcWorkbookName As String
    Dim srcSheetName As String

    Dim srcIdCol As Long, srcQtyCol As Long, srcBudgetCol As Long
    Dim destIdCol As Long, destQtyCol As Long, destBudgetCol As Long

    Dim srcLastRow As Long, destLastRow As Long
    Dim r As Long, logRow As Long

    Dim dict As Object
    Dim dictDup As Object
    Dim activityID As String
    Dim qtyVal As Variant, budgetVal As Variant

    Dim updatedCount As Long
    Dim missingCount As Long
    Dim duplicateCount As Long
    Dim blankIdCount As Long

    Dim logWs As Worksheet

    On Error GoTo ErrHandler

    Set destWb = ActiveWorkbook
    Set destWs = ActiveSheet

    srcWorkbookName = Trim$(InputBox( _
        "Enter the SOURCE workbook name exactly as it appears in Excel." & vbCrLf & _
        "Example: booook", _
        "Schedule Update", _
        "booook"))

    If srcWorkbookName = "" Then Exit Sub

    Set srcWb = FindOpenWorkbookByName(srcWorkbookName)

    If srcWb Is Nothing Then
        MsgBox "The source workbook is not open." & vbCrLf & _
               "Please open it first, then run the command again.", _
               vbExclamation, "Schedule Update"
        Exit Sub
    End If

    srcSheetName = Trim$(InputBox( _
        "Enter the source sheet name from workbook: " & srcWb.Name, _
        "Schedule Update", srcWb.Worksheets(1).Name))

    If srcSheetName = "" Then Exit Sub

    On Error Resume Next
    Set srcWs = srcWb.Worksheets(srcSheetName)
    On Error GoTo ErrHandler

    If srcWs Is Nothing Then
        MsgBox "Source sheet not found.", vbExclamation, "Schedule Update"
        Exit Sub
    End If

    srcIdCol = AskColumnByText(srcWs, _
        "Which column in the SOURCE sheet is Activity ID?" & vbCrLf & _
        "Example: A or 1 or Activity ID", "A")
    If srcIdCol = 0 Then Exit Sub

    srcQtyCol = AskColumnByText(srcWs, _
        "Which column in the SOURCE sheet is Total Quantity?" & vbCrLf & _
        "Example: F or 6 or Total Quantity", "F")
    If srcQtyCol = 0 Then Exit Sub

    srcBudgetCol = AskColumnByText(srcWs, _
        "Which column in the SOURCE sheet is Budgeted Amt?" & vbCrLf & _
        "Example: G or 7 or Budgeted Amt", "G")
    If srcBudgetCol = 0 Then Exit Sub

    destIdCol = AskColumnByText(destWs, _
        "Which column in THIS sheet is Activity ID?" & vbCrLf & _
        "Example: A or 1 or Activity ID", "A")
    If destIdCol = 0 Then Exit Sub

    destQtyCol = AskColumnByText(destWs, _
        "Which column in THIS sheet should receive Total Quantity?" & vbCrLf & _
        "Example: F or 6 or Total Quantity", "F")
    If destQtyCol = 0 Then Exit Sub

    destBudgetCol = AskColumnByText(destWs, _
        "Which column in THIS sheet should receive Budgeted Amt?" & vbCrLf & _
        "Example: G or 7 or Budgeted Amt", "G")
    If destBudgetCol = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set logWs = GetOrCreatePlannerLog(destWb)
    ClearPlannerLog logWs
    logRow = 2

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set dictDup = CreateObject("Scripting.Dictionary")
    dictDup.CompareMode = 1

    srcLastRow = srcWs.Cells(srcWs.Rows.Count, srcIdCol).End(xlUp).Row

    For r = 2 To srcLastRow
        activityID = Trim$(CStr(srcWs.Cells(r, srcIdCol).Value))

        If activityID <> "" Then
            qtyVal = srcWs.Cells(r, srcQtyCol).Value
            budgetVal = srcWs.Cells(r, srcBudgetCol).Value

            If dict.Exists(activityID) Then
                duplicateCount = duplicateCount + 1
                If Not dictDup.Exists(activityID) Then
                    dictDup.Add activityID, True
                    LogCategoryRow logWs, logRow, "Duplicate", r, activityID
                    logRow = logRow + 1
                End If
            Else
                dict.Add activityID, Array(qtyVal, budgetVal)
            End If
        End If
    Next r

    destLastRow = destWs.Cells(destWs.Rows.Count, destIdCol).End(xlUp).Row

    For r = 2 To destLastRow
        activityID = Trim$(CStr(destWs.Cells(r, destIdCol).Value))

        If activityID = "" Then
            blankIdCount = blankIdCount + 1
            LogCategoryRow logWs, logRow, "Blank", r, ""
            logRow = logRow + 1
        ElseIf dict.Exists(activityID) Then
            destWs.Cells(r, destQtyCol).Value = dict(activityID)(0)
            destWs.Cells(r, destBudgetCol).Value = dict(activityID)(1)
            updatedCount = updatedCount + 1
            LogCategoryRow logWs, logRow, "Updated", r, activityID
            logRow = logRow + 1
        Else
            missingCount = missingCount + 1
            LogCategoryRow logWs, logRow, "Missing", r, activityID
            logRow = logRow + 1
        End If
    Next r

    logWs.Visible = xlSheetVeryHidden

    MsgBox _
        "Schedule update completed." & vbCrLf & vbCrLf & _
        "Updated rows: " & updatedCount & vbCrLf & _
        "Missing Activities: " & missingCount & vbCrLf & _
        "Blank Activities: " & blankIdCount & vbCrLf & _
        "Duplicate Activities: " & duplicateCount, _
        vbInformation, "Schedule Update"

SafeExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbExclamation, "Planner"
    Resume SafeExit
End Sub

' =========================
' Review Highlighter
' =========================
Private Sub HighlightLoggedRows(ByVal categoryName As String, ByVal fillColor As Long)
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim targetRow As Long
    Dim lastCol As Long

    Set ws = ActiveSheet

    On Error Resume Next
    Set logWs = ActiveWorkbook.Worksheets(PLANNER_LOG_SHEET)
    On Error GoTo 0

    If logWs Is Nothing Then
        MsgBox "No Schedule Update log found. Run Sync Qty/Budget first.", vbExclamation, "Planner"
        Exit Sub
    End If

    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No logged rows found.", vbExclamation, "Planner"
        Exit Sub
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For r = 2 To lastRow
        If StrComp(CStr(logWs.Cells(r, 1).Value), categoryName, vbTextCompare) = 0 Then
            targetRow = CLng(Val(logWs.Cells(r, 2).Value))
            If targetRow > 0 Then
                ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastCol)).Interior.Color = fillColor
            End If
        End If
    Next r

    MsgBox categoryName & " highlighted.", vbInformation, "Planner"
End Sub

Private Sub ClearReviewColors()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim clr As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Application.ScreenUpdating = False

    For r = 1 To lastRow
        For c = 1 To lastCol
            clr = ws.Cells(r, c).Interior.Color

            If clr = RGB(255, 255, 0) _
               Or clr = RGB(255, 199, 206) _
               Or clr = RGB(255, 235, 156) _
               Or clr = RGB(198, 239, 206) Then
                ws.Cells(r, c).Interior.Pattern = xlNone
            End If
        Next c
    Next r

    Application.ScreenUpdating = True

    MsgBox "Only review colors cleared. WBS colors were kept.", vbInformation, "Planner"
End Sub

' =========================
' Helpers
' =========================
Private Function FindOpenWorkbookByName(ByVal baseName As String) As Workbook
    Dim wb As Workbook

    On Error Resume Next
    Set wb = Workbooks(baseName)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set FindOpenWorkbookByName = wb
        Exit Function
    End If

    On Error Resume Next
    Set wb = Workbooks(baseName & ".xlsx")
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set FindOpenWorkbookByName = wb
        Exit Function
    End If

    On Error Resume Next
    Set wb = Workbooks(baseName & ".xlsm")
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set FindOpenWorkbookByName = wb
        Exit Function
    End If
End Function

Private Function GetOrCreatePlannerLog(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(PLANNER_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = PLANNER_LOG_SHEET
    End If

    Set GetOrCreatePlannerLog = ws
End Function

Private Sub ClearPlannerLog(ByVal logWs As Worksheet)
    logWs.Cells.Clear
    logWs.Range("A1").Value = "Category"
    logWs.Range("B1").Value = "RowNumber"
    logWs.Range("C1").Value = "ActivityID"
End Sub

Private Sub LogCategoryRow(ByVal logWs As Worksheet, ByVal logRow As Long, ByVal categoryName As String, ByVal rowNumber As Long, ByVal activityID As String)
    logWs.Cells(logRow, 1).Value = categoryName
    logWs.Cells(logRow, 2).Value = rowNumber
    logWs.Cells(logRow, 3).Value = activityID
End Sub

Private Function AskColumnByText(ByVal ws As Worksheet, ByVal promptText As String, ByVal defaultValue As String) As Long
    Dim s As String
    Dim col As Long

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
    If AskColumnByText = 0 Then
        MsgBox "Column not found: " & s, vbExclamation, "Planner"
    End If
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), Trim$(headerText), vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Option Explicit

Public UserChoice As String

Private Sub chkDontShowAgain_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Labeld_Click()

End Sub

Private Sub lblInfo_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.UserChoice = ""
End Sub

Private Sub cmdContinue_Click()
    If chkDontShowAgain.Value = True Then
        SaveSetting "PlannerAddin", "Hints", "PlannedPercentage", "Hide"
    End If

    Me.UserChoice = "CONTINUE"
    Me.Hide
End Sub

Private Sub cmdBack_Click()
    Me.UserChoice = "BACK"
    Me.Hide
End Sub

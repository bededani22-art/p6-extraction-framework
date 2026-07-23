Option Explicit

Public UserChoice As String

Private Sub UserForm_Initialize()
    Me.UserChoice = ""
End Sub

Private Sub cmdContinue_Click()
    If chkDontShowAgain.Value = True Then
        SaveSetting "PlannerAddin", "Hints", "ActualPercentage", "Hide"
    End If

    Me.UserChoice = "CONTINUE"
    Me.Hide
End Sub

Private Sub cmdBack_Click()
    Me.UserChoice = "BACK"
    Me.Hide
End Sub

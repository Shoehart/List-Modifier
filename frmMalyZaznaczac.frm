Option Explicit

Private Sub CommandButton1_Click()
'Get the address from the RefEdit control.
'If frmMain.CheckBox4 = False Then
    frmMain.strRef2 = RefEdit1.Value
'End If
Unload Me
End Sub

'Private Sub UserForm_Initialize()
'    RefEdit1.Value = "Firmy!$D$2:$D$431737"
'End Sub
Private Sub UserForm_Click()

End Sub

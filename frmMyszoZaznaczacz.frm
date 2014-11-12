Option Explicit

Private Sub UserForm_Initialize()

'Dane przykładowe (zostawione TESTOWO, poprawić jak już wszystko zacznie działać!!!
If frmMain.CheckBox4.Value = True Then
    RefEdit1.Value = "Stanowiska!$A$2:$A$224469"
    RefEdit2.Value = "Sheet1!$A$2:$A$545"
    frmMain.strRef1 = RefEdit1.Value
    frmMain.strRef2 = RefEdit2.Value
End If

With frmMyszoZaznaczacz
    'Dane testowe
    RefEdit1.Value = "Sheet1!$A$2:$A$639"
    RefEdit2.Value = "Stanowiska!$A$2:$A$224469"
    
    'RefEdit1.Value = "Sheet1!$A$2:$A$10"
    'RefEdit2.Value = "Sheet1!$A$15:$A$21"
    
    frmMain.strRef1 = RefEdit1.Value
    frmMain.strRef2 = RefEdit2.Value

    'Skrócenie formularza w przypadku gdy dane wejściowe firm zostały już przygotowane
    If PDane_TN = True Then
        RefEdit1.Visible = False
        .Label1.Visible = False
        .Width = 180
        .RefEdit2.Left = 5
        .Label2.Left = 5
        .CommandButton1.Left = 130
    End If
End With
End Sub

Private Sub CommandButton1_Click()

'Get the address from the RefEdit control.
    If frmMain.CheckBox4 = False Then
        frmMain.strRef1 = RefEdit1.Value
        frmMain.strRef2 = RefEdit2.Value
    End If
   
Unload Me
frmMain.Show
End Sub

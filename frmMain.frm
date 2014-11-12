Option Explicit
Option Base 1
Public strRef1 As String, strRef2 As String

Private Sub CheckBox5_Click()
    If ListBox1.MultiSelect = fmMultiSelectMulti Then
        ListBox1.MultiSelect = fmMultiSelectSingle
    Else
        ListBox1.MultiSelect = fmMultiSelectMulti
    End If
End Sub

Private Sub CommandButton6_Click()
Dim zak As Range
Dim tempA As String, tempB As String
Dim i As Long, x As Long

Set zak = Worksheets(rRef(1, 2)).Range(rRef(2, 2))
ReDim tempArray(1 To zak.Rows.Count, 1 To zak.Columns.Count)

With frmMain
    .TextProgres.Visible = True
    .TextProgres.Font.Size = 50
    .ListBox1.Visible = False
End With

For i = 1 To UBound(tempArray, 1)
    If zak(i, 1).Value2 <> vbNullString Then
        tempA = zak(i, 1).Value2
        zak(i, 1).Value2 = WorksheetFunction.Trim(zak(i, 1).Value2)
        tempB = zak(i, 1).Value2
        If tempA <> tempB Then
            x = x + 1
        End If
    End If
    If i Mod 500 = 0 Then
        DoEvents ' Yield to operating system.
        With frmMain.TextProgres
            .Text = CStr(Format(i / UBound(zak.Value2, 1), "0%"))
        End With
    End If
Next i

frmMain.TextProgres.Visible = False
frmMain.ListBox1.Visible = True

MsgBox "zTRIMowałem: " & x & " rekordów", vbOKOnly
End Sub

Private Sub ListBox1_MouseMove( _
                        ByVal Button As Integer, ByVal Shift As Integer, _
                        ByVal x As Single, ByVal y As Single)
         HookListBoxScroll Me, Me.ListBox1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        UnhookListBoxScroll
End Sub

Private Sub UserForm_Initialize()
With Application
    .EnableEvents = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
    .StatusBar = False
End With

With frmMain
    .CheckBox2.Value = False
    .TextProgres.Visible = False
    .btnPickColour.Enabled = False
End With
    
    'Przypisanie danych wejściowych dla strRef 1 i 2
    strRef1 = "Sheet2!$A$2:$A$634" 'dane wyszukiwane w innym zakresie
    strRef2 = "YOUR_EEID_ORIG!$F$2:$F$856940" 'zakres danych do szukania w nim
    'Zmienne określające czy procedury przesły poprawnie przez dane
    PDane_TN = True 'Tak/Nie dla PrzygotujDane
    PDTablic_TN = False 'Tak/Nie dla PorównanieDwóchTablic
    
With Application
    .EnableEvents = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
End With
End Sub

Private Sub btnPickColour_Click()
    frmMain.tbColour.BackColor = PickNewColor
End Sub

Private Sub CheckBox2_Click()
    If frmMain.CheckBox2.Value = False Then
       btnPickColour.Enabled = False
    ElseIf CheckBox2.Value = True Then
       btnPickColour.Enabled = True
    End If
End Sub

Private Sub CheckBox4_Click()
    If frmMain.CheckBox4 = True Then
        If strRef1 = vbNullString Then
            strRef1 = "Sheet2!$I$1:$I$245"
            PDane_TN = True
        End If
        If strRef2 = vbNullString Then
            strRef2 = "Sheet2!$L$1:$O$12597"
            PDTablic_TN = True
        End If
    End If
End Sub

Private Sub CommandButton1_Click() ' OK Button - ma sprawdzić wszystkie CheckBoxy i wykonać odpowiednia obliczenia
Dim pomoc_a() As Variant
    
    If PDane_TN = False Then
        MsgBox "Przygotuj dane do prze przetwożenia!"
    End If
    
    If PDTablic_TN = False Then
        MsgBox "Zacznij od przycisku Zacznij Zabawę!"
    End If
    
    pomoc_a = frmMain.ListBox1.List
    If (CheckBox2 Or CheckBox1) And PDane_TN = PDTablic_TN = True Then
        KolorZakresu_i_Usuwanie TabelaPomocnicza:=pomoc_a
    End If
    Unload Me
End Sub

Private Sub CommandButton2_Click() ' "Zacznij zabawę" Button
Dim pomoc As Range
Dim tempArr() As String
Dim i As Long, j As Long

' Sprawdz czy wrzucić do konsoli dane przykładowe, czy faktycznie przeprowadzić obliczenia
If CheckBox4.Value = False Then
    Call DaneWejsciowe '- to jest własciwa procedura wywołująca przetważanie danych!
Else
    Set pomoc = Sheets("Sheet2").Range("$L$1:$N$309") 'once toutched Excel
    ReDim tempArr(1 To pomoc.Rows.Count, 1 To pomoc.Columns.Count)
    
    For i = 1 To UBound(tempArr, 1)
        For j = 1 To UBound(tempArr, 2)
            tempArr(i, j) = pomoc(i, j).Value2
        Next j
    Next i
    Call WrzucanieDoKonsoli(tempArr)
    PDane_TN = True
    PDTablic_TN = True
End If
End Sub

Private Sub CommandButton3_Click() ' Przygotuj Dane Button
    'frmMain.Hide
    frmMalyZaznaczacz.Show

    'If PDane_TN = False Then
        Call PrzygotujDane(1) 'po zabawie z samochodami zmienić na 2!
    'End If
End Sub

Private Sub CommandButton4_Click() ' Usuń Button
Dim i As Long, j As Long

If CheckBox5.Value = False Then
    'KASOWANIE w MultiSelect
    With Me.ListBox1
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                .RemoveItem (i)
            End If
        Next i
    frmMain.Caption = "Łącznie: " & .ListCount & " pozycji."
    End With
Else
    'KASOWANIE w Single Select
    Dim arrPomoc(), tempPomoc As String
    With frmMain.ListBox1
         If ListBox1.ListIndex <> -1 Then
            i = ListBox1.ListIndex
            j = 1
            tempPomoc = .List(i, j)
                Do
                    .RemoveItem i
                    frmMain.Caption = "Łącznie: " & .ListCount & " pozycji."
                If i >= .ListCount Then
                    Exit Do
                End If
                Loop While .List(i, j) = tempPomoc
          End If
    End With
End If
End Sub
Private Sub CommandButton5_Click()
Dim rng As Range
Dim tempArray() As Variant

frmMalyZaznaczacz.Show
frmMalyZaznaczacz.Caption = "Wskaż gdzie zdumpować dane?"

tempArray = ListBox1.List
With Sheets(aMain.rRef(1, 2))
    '.Columns("A:A").NumberFormat = "@"
    Set rng = .Range(rRef(2, 2))
    Set rng = rng.Resize(UBound(tempArray, 1), UBound(tempArray, 2) + 1)
    rng.ClearContents
    rng = tempArray
End With
End Sub

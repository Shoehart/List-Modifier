Option Base 1
Option Explicit

'Enables mouse wheel scrolling in controls
#If Win64 Then
    Private Type POINTAPI
       XY As LongLong
    End Type
#Else
    Private Type POINTAPI
           x As Long
           y As Long
    End Type
#End If

Private Type MOUSEHOOKSTRUCT
    Pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long ' not sure if this should be LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" _
                                            Alias "GetWindowLongPtrA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #End If
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As LongPtr, _
                                                            ByVal hmod As LongPtr, _
                                                            ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As LongPtr, _
                                                           lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr) As LongPtr ' MAYBE Long
    #If Win64 Then
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal Point As LongLong) As LongPtr    '
    #Else
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As LongPtr    '
    #End If
    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
                                                            ByRef lpPoint As POINTAPI) As LongPtr   'MAYBE Long
#Else
    Private Declare Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As Long, _
                                                            ByVal nIndex As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As Long, _
                                                            ByVal hmod As Long, _
                                                            ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As Long, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As Long, _
                                                           lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                            ByRef lpPoint As POINTAPI) As Long
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
Dim N As Long
Private mCtl As Object
Private mbHook As Boolean
#If VBA7 Then
    Private mLngMouseHook As LongPtr
    Private mListBoxHwnd As LongPtr
#Else
    Private mLngMouseHook As Long
    Private mListBoxHwnd As Long
#End If

Public PDane_TN As Boolean, PDTablic_TN As Boolean

'================================================================================
' Sub DaneWejsciowe
'
' przekazuje do innych funkcji zakresy danych wraz z nazwami arkuszy
'================================================================================
Sub DaneWejsciowe()

' Sprawdz czy wrzucić dane przykładowe/już przetworzone czy przeprowadzić faktyczną pracę
' na żywych danych z podanych zakresów.

If frmMain.strRef1 = vbNullString Or frmMain.strRef2 = vbNullString = True Then
    'frmMain.Hide
    frmMyszoZaznaczacz.Show
    Unload frmMyszoZaznaczacz
End If

If PDane_TN = False Then
    Call PrzygotujDane(3)
End If

If PDTablic_TN = False Then
    PorwnanieDwochTablic
End If

End Sub

'===================================================================================
' Sub PrzygotujDane(nrKolumny As Long)
'
' Służy przede wszystkim do czyszczenia nazw firm!
'
' Czyści dane wejściowe z pustych rekordów, zbędnych znaków
' Usuwa wszystkie rekordy krótsze jak 2 znaki i dłuższe jak 30 oraz usuwa duplikaty.
'===================================================================================
Sub PrzygotujDane(nrKolumny As Long)
Dim arrData() As Variant, arrData2() As String, tempArray() As Variant
Dim i As Long, j As Long, lCount As Long, LR As Long, pomc As Long, N As Long
Dim oRegex As Object
Dim r As Variant
Dim rng As Range, zak As Range

'Zgranie danych z arkusza
Set zak = Worksheets(rRef(1, 2)).Range(rRef(2, 2))
arrData = zak.Value2

'Wywalanie duplikatów z danych wejściowych
UsuwanieDuplikatow DaneWejsciowe:=arrData

ReDim arrData2(2, 1 To zak.Rows.Count)
N = 0
lCount = UBound(arrData, 1)

'schowanie ListBoxa (czyli konsoli) i pokazanie % postępu pracy
With frmMain
    .TextProgres.Visible = True
    .TextProgres.Font.Size = 50
    .ListBox1.Visible = False
End With


For i = 1 To UBound(arrData, 1)
    If oRegex Is Nothing Then Set oRegex = CreateObject("vbscript.regexp")
    
    'czyszczenie stringów ze zbędnych znaków z wykorzystanie RegEx'a oraz ich TRIM
    With oRegex
        .Global = True
        .Pattern = "[^A-Za-z0-9]" 'Allow A-Z, a-z, 0-9, and space
        If Len(arrData(i, 1)) < 3 Then
            arrData(i, 1) = vbNullString
        Else
         ' wszystkie lowercase, i wyczyszczone ze zbędnych znaków i spacji
          'arrData(i, 1) = LCase(arrData(i, 1))
          arrData(i, 1) = .Replace(arrData(i, 1), " ")
          arrData(i, 1) = WorksheetFunction.Trim(arrData(i, 1))
        End If
    'koniec czyszczenia RegEx'em i TRIMu
        
    'kod na wydzialanie poszczególnych słów.
        r = Split(arrData(i, 1), " ")
        LR = UBound(r)
        If LR >= 0 Then
         Select Case LR
            Case Is = 0
                pomc = pomc + 1
            Case Is > 0
                pomc = LR + pomc + 1
         End Select
         'ReDim Preserve arrData2(2, 1 To pomc)
            For j = 0 To LR
              ' wywalanie wszystkich strinógw krótszych jak 1 znak i dłuższych jak 30 znaków
                If Len(r(j)) <= 1 Or Len(r(j)) >= 30 Then
                    arrData2(1, j + 1 + N) = vbNullString
                    pomc = pomc - 1
                Else
                    arrData2(1, j + 1 + N) = r(j)
                    arrData2(2, j + 1 + N) = arrData(i, 1)
                End If
                
                If i Mod 100 = 0 Then
                    'Application.StatusBar = "Wykonałem " & Format(i * j / lCount, "0%") & " kalkulacji."
                    DoEvents ' Yield to operating system.
                    With frmMain.TextProgres
                        .Text = CStr(Format(i / lCount, "0%"))
                    End With
                End If
                
            Next j
            N = pomc
        End If
    'koniec kodu ekstrahowania poszczególnych słów
    End With
Next i
ReDim Preserve arrData2(2, 1 To N)
Set oRegex = Nothing

'pokazanie listboxa i schowanie countera postępu pracy
With frmMain
    .TextProgres.Visible = False
    .ListBox1.Visible = True
End With

ReDim tempArray(1 To UBound(arrData2, 2), 1 To UBound(arrData2, 1))
'obrocenie arrData2 do wersji pionowej/wertykalnej
For i = 1 To UBound(arrData2, 1)
    For j = 1 To UBound(arrData2, 2)
      tempArray(j, i) = arrData2(i, j)
    Next j
Next i
Erase arrData2

'usuwanie duplikatów i sortowanie wynikowej tablicy
UsuwanieDuplikatow DaneWejsciowe:=tempArray
Call QuickSortArray(tempArray, , , nrKolumny) ' sortowanie na tablicy horyzontalnej nie działa
'koniec sprawdzenia
   
' Wrzucenie danych do konsoli
Call WrzucanieDoKonsoli(tempArray)

'Czy Sub był już execowany (True = TAK, False = NIE)
PDane_TN = True
End Sub

'================================================================================
' Sub PorwnanieDwochTablic
'
' Znajduje dane z malej tabeli (zak_M) w dużej (zak_D) i daje możliwość
' ich zaznaczenia/skasowania w danych wyjściowych, bądź wrzucenia do konsoli.
'================================================================================
Sub PorwnanieDwochTablic()
Dim i As Long, j As Long, LR As Long, lCount As Long
Dim mala() As Variant, duza() As Variant, pomoc() As Variant, pomoc_a() As Variant, r As Variant
Dim temp_mala As String, temp_duza As String
Dim zak_M As Range, zak_D As Range
    
'If PDane_TN = False Then
    Set zak_M = Sheets(rRef(1, 1)).Range(rRef(2, 1)) 'once toutched Excel
    ReDim mala(1 To zak_M.Rows.Count)
    mala = zak_M.Value2
'Else
'    mala = frmMain.ListBox1.List
'End If
    
    Set zak_D = Sheets(rRef(1, 2)).Range(rRef(2, 2)) 'once toutched Excel
    ReDim duza(1 To zak_D.Rows.Count)
    ReDim pomoc(1 To zak_D.Rows.Count, 1 To 3)
    duza = zak_D.Value2 'znaznacza zakres z aktywnego arkusza, a nie z zakresu chwyconego wczesniej
    
    LR = 1
    lCount = UBound(duza, 1) * UBound(mala, 1)
    
    With frmMain
        .TextProgres.Visible = True
        .TextProgres.Font.Size = 50
        .ListBox1.Visible = False
    End With
    
    'Wyszukiwanie takich samych z Małej w Dużej i zaznaczanie ich miejsca w zmiennej
    ' pomocniczej "pomoc"
        For i = 1 To UBound(duza, 1)
         temp_duza = duza(i, 1)
          If temp_duza <> vbNullString Then
            For j = 1 To UBound(mala, 1)
              temp_mala = mala(j, 1)
                'porównanie poszczególnych słów ze stanowiska z listą podejrzanych nazwa firm
                    If PorownajCiagiZnakow(temp_mala, temp_duza, True, True) = True Then
                       pomoc(LR, 1) = i
                       pomoc(LR, 2) = temp_mala
                       pomoc(LR, 3) = temp_duza
                       LR = LR + 1
                       'Exit For
                    End If
            Next j
          End If
        'Update StatusBaru o % postępu prac
        If i Mod 500 = 0 Then
            'Application.StatusBar = "Wykonałem " & Format(i * j / lCount, "0%") & " kalkulacji."
            DoEvents ' Yield to operating system.
            With frmMain.TextProgres
                .Text = CStr(Format(i * j / lCount, "0%"))
            End With
        End If
        Next i
        
    'give statusbar back to Excel
    'Application.StatusBar = False
    frmMain.TextProgres.Visible = False
    frmMain.ListBox1.Visible = True

    'Transpozycja tablicy pomoc() z horyzontalnej na wertykalną.
    ReDim pomoc_a(1 To LR, 1 To 3)
    For i = 1 To LR
     For j = 1 To 3
        pomoc_a(i, j) = pomoc(i, j)
     Next j
    Next i
    
    Erase pomoc, duza, mala
    
    'Posortowanie tablicy z wynikami po wartościach z "Małej"
    Call QuickSortArray(pomoc_a, , , 2)
    Call WrzucanieDoKonsoli(pomoc_a)
    
    PDTablic_TN = True
End Sub
'================================================================================
' Sub KolorZakresu_i_Usuwanie
'
' Kolorowanie znalezionych stringów z Małej w podanym zakresie z opcją usunięcia
' podejrzanych danych ze stringów.
'================================================================================
Sub KolorZakresu_i_Usuwanie(ByVal TabelaPomocnicza As Variant)
Dim i As Long, iR As Long, iG As Long, iB As Long, lRGB As Long
Dim zak_D As Range
Dim r As Long
Dim CzyUsunac As Boolean, CzyKolorowac As Boolean

Set zak_D = Sheets(rRef(1, 2)).Range(rRef(2, 2))
Call QuickSortArray(TabelaPomocnicza, , , 0)

CzyUsunac = frmMain.CheckBox1.Value
CzyKolorowac = frmMain.CheckBox2.Value
   
'kolorowanie kolorem pobranym z frmMain i zmiana koloru OEM na RGB
      lRGB = Abs(frmMain.tbColour.BackColor)
      iB = (lRGB \ 65536) And &HFF
      iG = (lRGB \ 256) And &HFF
      iR = lRGB And &HFF
      
      For i = 0 To UBound(TabelaPomocnicza, 1)
      r = TabelaPomocnicza(i, 0)
        If CzyKolorowac Then
            zak_D(r).Interior.Color = RGB(iR, iG, iB)
        End If
        If CzyUsunac Then
            zak_D(r).Value = Replace(zak_D(r).Value, TabelaPomocnicza(i, 1), vbNullString, , , vbTextCompare)
            zak_D(r).Value = WorksheetFunction.Trim(zak_D(r))
        End If
      Next i
End Sub

'================================================================================
' Function SetColumnWidths(ctrList As Variant)
'
' Funkcja ustawia szerokości kolumn w konsoli w zależnosci od ilości dostarczonych
' danych. Dostosowuje .ColumnWidths do najdłuższego ciągu znaków w dane kolumnie
' z dostarczonych danych.
'================================================================================

Public Function SetColumnWidths(ctrList As Variant) As String
Dim i As Long, x As Long, y As Long, z As Long, pomoc As Long, colNum As Long, _
    rowNum As Long, listBoxWidths As Long, sumOfLength As Long
Dim aryLen() As Single, colRatio As Single
Dim ctrValue As String, colWidths As String

' Ilość kolumn i wierszy w przekazanej tablicy
colNum = UBound(ctrList, 2)
rowNum = UBound(ctrList, 1)

' Jeżeli tablica ma tylko jedną kolumnę to kończy funkcję
If colNum = 1 Then
    SetColumnWidths = listBoxWidths - 15
    Exit Function
End If

    'Tablica z szerokościami dla poszczególnych kolumn
    ReDim aryLen(colNum)
    
    ' Sprawdzenie wszystkich wartości w kolumnach w poszukiwaniu najdłuższego stringa
    For x = 1 To colNum ' dla każdej kolumny w tablicy
        For y = 1 To rowNum ' dla każdego wiersza w tablicy
            ctrValue = ctrList(y, x)
            If Len(ctrValue) > aryLen(x) Then
                aryLen(x) = Len(ctrValue)
            End If
        Next y
        sumOfLength = sumOfLength + aryLen(x)
    Next x
    
    ' Szerokość okienka konsoli i współczynnik zmiany ilości znaków na px
    listBoxWidths = frmMain.ListBox1.Width
    colRatio = 2.37
    
    ' Dla każdej kolumny procentowa szerokość w odniesieniu do szerokości okienka i sumy długości wszystkich znaków
    For i = 1 To UBound(aryLen) - 1
        If aryLen(i) > 1 Then
            pomoc = Round((aryLen(i) * colRatio) / sumOfLength * listBoxWidths, 0)
            z = aryLen(i)
            If pomoc < 100 Then
                aryLen(i) = pomoc
            Else
                aryLen(i) = 100
            End If
            listBoxWidths = listBoxWidths - aryLen(i)
            sumOfLength = sumOfLength - z
        Else
            aryLen(i) = 0
        End If
    Next i
    
    ' Przygotowanie stringa dla .ColumnWidth
        For i = 1 To colNum - 1 'For each stored maximum lenght
            colWidths = colWidths & (aryLen(i) & ";")
        Next i
        colWidths = colWidths & "200"
    
    ' Zwrócenie wartości w stringu dla .ColumnWidth
    SetColumnWidths = colWidths
End Function

'================================================================================
' Sub WrzucanieDoKonsoli
'
' Wrzuca dane do konsoli po wcześniejszym dostosowaniu .ColumnWidth przy pomocy
' funkcji SetColumnWidths()
'================================================================================
Sub WrzucanieDoKonsoli(ByVal tempArray As Variant)
Dim LR As Long, i As Long, nrRow As Long, nrCol As Long

nrCol = UBound(tempArray, 2)
nrRow = UBound(tempArray, 1)

With frmMain.ListBox1
    .ColumnCount = nrCol
    .ColumnWidths = SetColumnWidths(tempArray)
    .List = tempArray
    frmMain.Caption = "Łącznie: " & .ListCount - 1 & " wystąpień."
End With

End Sub

'================================================================================
' Function PorownajCiagiZnakow
'
' Porównuje dwa stringi ze sobą. Funkcja pozwala znaleźć wysąpienie stringa StringToBeFound
' w StringToSearchIn. Pozwa a ustawić czy ma szukać "całych słów", jak również
' czy ma rozpoznawać wielkość liter.
'================================================================================

Function PorownajCiagiZnakow( _
         ByVal StringToBeFound As String, _
         ByVal StringToSearchIn As String, _
         WholeWord As Boolean, _
Optional CaseSens As Boolean = False) As Boolean
              
Dim i As Long
Dim r As Variant
      
  ' default return value if value not found in array
    PorownajCiagiZnakow = False
    
  ' Sprawdzenie czy jest CaseSense włączony
    If CaseSens = True Then
        StringToSearchIn = LCase(StringToSearchIn)
        StringToBeFound = LCase(StringToBeFound)
    End If
    
    If InStr(StringToSearchIn, StringToBeFound) > 0 Then
        'Sprawdzenie czy sprawdzać tylko całe słowa
        If WholeWord = True Then
            r = Split(StringToSearchIn, " ")
            For i = 0 To UBound(r)
                If StrComp(StringToBeFound, r(i), vbBinaryCompare) = 0 Then
                    PorownajCiagiZnakow = True
                    Exit For
                End If
            Next i
        Else
            PorownajCiagiZnakow = True
        End If
    End If
End Function

'================================================================================
' PickNewColor
'
' Excelowa funkcja wyboru koloru
'================================================================================

Function PickNewColor(Optional i_OldColor As Double = xlNone) As Double
Const BGColor As Long = 13160660  'background color of dialogue
Const ColorIndexLast As Long = 32 'index of last custom color in palette

Dim myOrgColor As Double          'original color of color index 32
Dim myNewColor As Double          'color that was picked in the dialogue
Dim myRGB_R As Integer            'RGB values of the color that will be
Dim myRGB_G As Integer            'displayed in the dialogue as
Dim myRGB_B As Integer            '"Current" color
  
  'save original palette color, because we don't really want to change it
  myOrgColor = ActiveWorkbook.Colors(ColorIndexLast)
  
  If i_OldColor = xlNone Then
    'get RGB values of background color, so the "Current" color looks empty
    Kolor2RGB BGColor, myRGB_R, myRGB_G, myRGB_B
  Else
    'get RGB values of i_OldColor
    Kolor2RGB i_OldColor, myRGB_R, myRGB_G, myRGB_B
  End If
  
  'call the color picker dialogue
  If Application.Dialogs(xlDialogEditColor).Show(ColorIndexLast, _
     myRGB_R, myRGB_G, myRGB_B) = True Then
    '"OK" was pressed, so Excel automatically changed the palette
    'read the new color from the palette
    PickNewColor = ActiveWorkbook.Colors(ColorIndexLast)
    'reset palette color to its original value
    ActiveWorkbook.Colors(ColorIndexLast) = myOrgColor
  Else
    '"Cancel" was pressed, palette wasn't changed
    'return old color (or xlNone if no color was passed to the function)
    PickNewColor = i_OldColor
  End If
End Function

'================================================================================
' Kolor2RGB
'
' Zmiana koloru z OEM na składowe RGB (R, G, B)
'================================================================================
Sub Kolor2RGB(ByVal i_Color As Long, o_R As Integer, o_G As Integer, o_B As Integer)
  o_R = i_Color Mod 256
  i_Color = i_Color \ 256
  o_G = i_Color Mod 256
  i_Color = i_Color \ 256
  o_B = i_Color Mod 256
End Sub

'================================================================================
' Sub UsuwanieDuplikow()
'
' Usuwa duplikaty z tablic 2D i z wertków z wykorzystaniem Classy Dictionary.
'================================================================================
Function UsuwanieDuplikatow(ByRef DaneWejsciowe As Variant)
Dim arr As Dictionary
Dim v As Variant, r(1 To 2) As Variant
Dim i As Long, nrCol As Long, nrRow As Long, nNr As Long, nMultiCol As Long, N As Byte, j As Long
Dim varArrFromDict As Variant

Set arr = CreateObject("Scripting.Dictionary")
N = NumberOfArrayDimensions(DaneWejsciowe)

If N = 2 Then
    'Sprawdzenie czy wejsciowa tablica jest horyzontalna czy wertykalna
    nrRow = UBound(DaneWejsciowe, 1)
    nrCol = UBound(DaneWejsciowe, 2)
    If nrRow > nrCol Then
        'tablica horyzontalna
        nNr = 1
        If nrCol > 1 Then
            nMultiCol = 1
        Else
            nMultiCol = 0
        End If
    Else
        'tablica wertykalna
        nNr = 2
        If nrRow >= 1 Then
            nMultiCol = 1
        Else
            nMultiCol = 0
        End If
    End If
ElseIf N = 1 Then
        nNr = 1
        nrRow = UBound(DaneWejsciowe)
        nMultiCol = 0
Else
    MsgBox "Daj mi coś do pracy, jakąkolwiek tabelkę!"
End If

If nMultiCol = 1 Then
    Set arr = New Dictionary
    arr.CompareMode = TextCompare
    
    For v = 1 To UBound(DaneWejsciowe, nNr)
        If nNr = 1 Then
            r(1) = DaneWejsciowe(v, 1)
            r(2) = DaneWejsciowe(v, 2)
        Else
            r(1) = DaneWejsciowe(1, v)
            r(2) = DaneWejsciowe(2, v)
        End If
        If arr.Exists(r(1)) <> True Then
            arr.Add r(1), Array(r)
        End If
    Next v
Else
    Set arr = New Dictionary
    For Each v In DaneWejsciowe
        If v <> vbNullString Then
            If arr.Exists(v) <> True Then
                arr.Add v, v
            End If
        End If
    Next v
End If

ReDim DaneWejsciowe(1 To arr.Count, 1 To 2)

For i = 0 To arr.Count - 1
varArrFromDict = arr.Items(i)
    If IsArray(varArrFromDict) Then
        DaneWejsciowe(i + 1, 1) = varArrFromDict(1)(1)
        DaneWejsciowe(i + 1, 2) = varArrFromDict(1)(2)
    Else
        DaneWejsciowe(i + 1, 1) = varArrFromDict
    End If
Next i

varArrFromDict = vbNullString
Erase r
Set arr = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' rRef(ktoryCzlon As Byte, ktoryElement As Byte) As String
'
' Funkcja rozbija adres zaznaczonych MyszoZaznaczaczem danych na nazwę arkusza
' i adres danych pobierając wartości z Classy clsGroupRef
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function rRef(ktoryCzlon As Byte, ktoryElement As Byte) As String

'wybiera, który element ze stringa wyciągnąć
If ktoryElement = 1 Then 'wybiera miedzy RefEdit1 a RefEdit2
    Select Case ktoryCzlon
        Case Is = 1 'nazwa arkusza 1
            rRef = Left(frmMain.strRef1, InStr(frmMain.strRef1, "!") - 1)
        Case Is = 2 'zakres danych 1
            rRef = Right(frmMain.strRef1, Len(frmMain.strRef1) - InStr(frmMain.strRef1, "!"))
    End Select
Else
    Select Case ktoryCzlon
        Case Is = 1 'nazwa arkusza 2
            rRef = Left(frmMain.strRef2, InStr(frmMain.strRef2, "!") - 1)
        Case Is = 2 'zakres danych 2
            rRef = Right(frmMain.strRef2, Len(frmMain.strRef2) - InStr(frmMain.strRef2, "!"))
    End Select
End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MOUSE Scroll
'
' Scrolowanie danych w ListBoxie, kod wzięty z neta (https://social.msdn.microsoft.com/Forums/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub HookListBoxScroll(frm As Object, ctl As Object)
    Dim tPT As POINTAPI
    #If VBA7 Then
        Dim lngAppInst As LongPtr
        Dim hwndUnderCursor As LongPtr
    #Else
        Dim lngAppInst As Long
        Dim hwndUnderCursor As Long
    #End If
    GetCursorPos tPT
    #If Win64 Then
        hwndUnderCursor = WindowFromPoint(tPT.XY)
    #Else
        hwndUnderCursor = WindowFromPoint(tPT.x, tPT.y)
    #End If
    If TypeOf ctl Is UserForm Then
        If Not frm Is ctl Then
               ctl.SetFocus
        End If
    Else
        If Not frm.ActiveControl Is ctl Then
             ctl.SetFocus
        End If
    End If
    If mListBoxHwnd <> hwndUnderCursor Then
        UnhookListBoxScroll
        Set mCtl = ctl
        mListBoxHwnd = hwndUnderCursor
        #If Win64 Then
            lngAppInst = GetWindowLongPtr(mListBoxHwnd, GWL_HINSTANCE)
        #Else
            lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        #End If
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx( _
                                            WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
        End If
    End If
End Sub

Sub UnhookListBoxScroll()
    If mbHook Then
        Set mCtl = Nothing
        UnhookWindowsHookEx mLngMouseHook
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub
#If VBA7 Then
    Private Function MouseProc( _
                            ByVal nCode As Long, ByVal wParam As Long, _
                            ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
        Dim idx As Long
        On Error GoTo errH
        If (nCode = HC_ACTION) Then
            #If Win64 Then
                If WindowFromPoint(lParam.Pt.XY) = mListBoxHwnd Then
                    If wParam = WM_MOUSEWHEEL Then
                        MouseProc = True
                        If TypeOf mCtl Is Frame Then
                            If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                            idx = idx + mCtl.ScrollTop
                            If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                                mCtl.ScrollTop = idx
                            End If
                        ElseIf TypeOf mCtl Is UserForm Then
                            If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                            idx = idx + mCtl.ScrollTop
                            If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                                mCtl.ScrollTop = idx
                            End If
                        Else
                            If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                            idx = idx + mCtl.ListIndex
                            If idx >= 0 Then mCtl.ListIndex = idx
                        End If
                    Exit Function
                    End If
                Else
                    UnhookListBoxScroll
                End If
            #Else
                If WindowFromPoint(lParam.Pt.x, lParam.Pt.y) = mListBoxHwnd Then
                    If wParam = WM_MOUSEWHEEL Then
                        MouseProc = True
                        If TypeOf mCtl Is Frame Then
                            If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                            idx = idx + mCtl.ScrollTop
                            If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                                mCtl.ScrollTop = idx
                            End If
                        ElseIf TypeOf mCtl Is UserForm Then
                            If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                            idx = idx + mCtl.ScrollTop
                            If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                                mCtl.ScrollTop = idx
                            End If
                        Else
                            If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                            idx = idx + mCtl.ListIndex
                            If idx >= 0 Then mCtl.ListIndex = idx
                        End If
                        Exit Function
                    End If
                Else
                    UnhookListBoxScroll
                End If
            #End If
        End If
        MouseProc = CallNextHookEx( _
                                mLngMouseHook, nCode, wParam, ByVal lParam)
        Exit Function
errH:
        UnhookListBoxScroll
    End Function
#Else
    Private Function MouseProc( _
                            ByVal nCode As Long, ByVal wParam As Long, _
                            ByRef lParam As MOUSEHOOKSTRUCT) As Long
        Dim idx As Long
        On Error GoTo errH
        If (nCode = HC_ACTION) Then
            If WindowFromPoint(lParam.Pt.x, lParam.Pt.y) = mListBoxHwnd Then
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
                    If TypeOf mCtl Is Frame Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is UserForm Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    Else
                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                        idx = idx + mCtl.ListIndex
                        If idx >= 0 Then mCtl.ListIndex = idx
                    End If
                    Exit Function
                End If
            Else
                UnhookListBoxScroll
            End If
        End If
        MouseProc = CallNextHookEx( _
        mLngMouseHook, nCode, wParam, ByVal lParam)
        Exit Function
errH:
        UnhookListBoxScroll
    End Function
#End If

Public Function SelectAll(lst As ListBox) As Boolean
On Error GoTo Err_Handler
    'Purpose:   Select all items in the multi-select list box.
    'Return:    True if successful
    'Author:    Allen Browne. http://allenbrowne.com  June, 2006.
    Dim lngRow As Long

    If lst.MultiSelect Then
        For lngRow = 0 To lst.ListCount - 1
            lst.Selected(lngRow) = True
        Next
        SelectAll = True
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    Call LogError(Err.Number, Err.Description, "SelectAll()")
    Resume Exit_Handler
End Function

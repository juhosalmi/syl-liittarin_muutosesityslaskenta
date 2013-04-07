Attribute VB_Name = "Main"
Sub Main()
    
    ' Lasketaan kaikki esitykset ja kannatukset yhteensä
    ' Call Kaikki
    
    Application.DisplayAlerts = False
    
    ' AYY
    Dim AYY As CPresenter
    Set AYY = New CPresenter
    Call AYY.Init("(AYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(AYY)
    
    ' HYY
    Dim HYY As CPresenter
    Set HYY = New CPresenter
    Call HYY.Init("(HYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(HYY)

    ' ISYY
    Dim ISYY As CPresenter
    Set ISYY = New CPresenter
    Call ISYY.Init("(ISYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(ISYY)

    ' JYY
    Dim JYY As CPresenter
    Set JYY = New CPresenter
    Call JYY.Init("(JYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(JYY)

    ' KUVYO
    Dim KUVYO As CPresenter
    Set KUVYO = New CPresenter
    Call KUVYO.Init("(KUVYO)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(KUVYO)

    ' LYY
    Dim LYY As CPresenter
    Set LYY = New CPresenter
    Call LYY.Init("(LYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(LYY)

    ' LTKY
    Dim LTKY As CPresenter
    Set LTKY = New CPresenter
    Call LTKY.Init("(LTKY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(LTKY)

    ' OYY
    Dim OYY As CPresenter
    Set OYY = New CPresenter
    Call OYY.Init("(OYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(OYY)

    ' SAY
    Dim SAY As CPresenter
    Set SAY = New CPresenter
    Call SAY.Init("(SAY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(SAY)

    ' SHS
    Dim SHS As CPresenter
    Set SHS = New CPresenter
    Call SHS.Init("(SHS)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(SHS)

    ' TTYY
    Dim TTYY As CPresenter
    Set TTYY = New CPresenter
    Call TTYY.Init("(TTYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(TTYY)

    ' Tamy
    Dim Tamy As CPresenter
    Set Tamy = New CPresenter
    Call Tamy.Init("(Tamy)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(Tamy)

    ' TeYO
    Dim TeYO As CPresenter
    Set TeYO = New CPresenter
    Call TeYO.Init("(TeYO)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(TeYO)

    ' TYY
    Dim TYY As CPresenter
    Set TYY = New CPresenter
    Call TYY.Init("(TYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(TYY)

    ' VYY
    Dim VYY As CPresenter
    Set VYY = New CPresenter
    Call VYY.Init("(VYY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(VYY)

    ' ÅAS
    Dim ÅAS As CPresenter
    Set ÅAS = New CPresenter
    Call ÅAS.Init("(ÅAS)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(ÅAS)

    ' SKY
    Dim SKY As CPresenter
    Set SKY = New CPresenter
    Call SKY.Init("(SKY)", Sheets("MUUTOSESITYSEXCEL"))
    Call CreateWorksheet(SKY)
    
End Sub

Function CreateWorksheet(presenter As CPresenter) As Worksheet
    
    Sheets(presenter.Nimi).Delete ' Comment this if there isn't such sheet
    
    Dim Ws As Worksheet
    Set Ws = Sheets.Add
    
    Ws.Name = presenter.Nimi
    
    Ws.Cells(1, 1) = presenter.Nimi
    
    Ws.Cells(2, 1) = "Esitykset"
    Ws.Cells(3, 1) = "Esityksiä läpi"
    Ws.Cells(4, 1) = "Esityksiä läpi muutoksin"
    Ws.Cells(5, 1) = "Läpäisy"
    
    Ws.Range("N2:N4").Formula = "=SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])"
    Ws.Range("B5, D5, F5, H5, J5, L5, N5").Formula = "=SUM(R[-2]C:R[-1]C)/R[-3]C"
    Ws.Range("B5, D5, F5, H5, J5, L5, N5").NumberFormat = "0.00%"
    
    Ws.Cells(7, 1) = "Kannatukset"
    Ws.Cells(8, 1) = "Kannatuksia läpi"
    Ws.Cells(9, 1) = "Kannatuksia läpi muutoksin"
    Ws.Cells(10, 1) = "Läpäisy"
    
    Ws.Range("N7:N9").Formula = "=SUM(RC[-12],RC[-10],RC[-8],RC[-6],RC[-4],RC[-2])"
    Ws.Range("B10, D10, F10, H10, J10, L10, N10").Formula = "=SUM(R[-2]C:R[-1]C)/R[-3]C"
    Ws.Range("B10, D10, F10, H10, J10, L10, N10").NumberFormat = "0.00%"
    
    Ws.Cells(1, 2) = "Lipa"
    Ws.Cells(1, 4) = "Tosu"
    Ws.Cells(1, 6) = "Talousarvio"
    Ws.Cells(1, 8) = "Kannanotot"
    Ws.Cells(1, 10) = "Yhteiskannanotto"
    Ws.Cells(1, 12) = "Ponnet"
    Ws.Cells(1, 14) = "Yht."
    
    Dim i As Integer
    For i = 1 To presenter.Kokouskohdat.Count
        Ws.Cells(2, 2 * i) = presenter.Kokouskohdat(i).esitykset
        Ws.Cells(3, 2 * i) = presenter.Kokouskohdat(i).esityksetLapi
        Ws.Cells(4, 2 * i) = presenter.Kokouskohdat(i).esityksetLapiMuutoksin
        
        Ws.Cells(2, 2 * i + 1).Formula = "=RC[-1]/Kaikki!RC[-" & i & "]"
        Ws.Range(Ws.Cells(3, 2 * i + 1), Ws.Cells(4, 2 * i + 1)).Merge
        Ws.Cells(3, 2 * i + 1).Formula = "=(RC[-1]+R[+1]C[-1])/(Kaikki!RC[-" & i & "]+Kaikki!R[+1]C[-" & i & "])"
        
        'Ws.Cells(3, 2 * i + 1).Formula = "=RC[-1]/Kaikki!RC[-" & i & "]"
        'Ws.Cells(4, 2 * i + 1).Formula = "=RC[-1]/Kaikki!RC[-" & i & "]"
        
        Ws.Cells(7, 2 * i) = presenter.Kokouskohdat(i).kannatukset
        Ws.Cells(8, 2 * i) = presenter.Kokouskohdat(i).kannatuksetLapi
        Ws.Cells(9, 2 * i) = presenter.Kokouskohdat(i).kannatuksetLapiMuutoksin
        
        Ws.Cells(7, 2 * i + 1).Formula = "=RC[-1]/Kaikki!R[-5]C[-" & i & "]"
        Ws.Range(Ws.Cells(8, 2 * i + 1), Ws.Cells(9, 2 * i + 1)).Merge
        Ws.Cells(8, 2 * i + 1).Formula = "=(RC[-1]+R[+1]C[-1])/(Kaikki!R[-5]C[-" & i & "]+Kaikki!R[-4]C[-" & i & "])"
        
        'Ws.Cells(8, 2 * i + 1).Formula = "=RC[-1]/Kaikki!R[-5]C[-" & i & "]"
        'Ws.Cells(9, 2 * i + 1).Formula = "=RC[-1]/Kaikki!R[-5]C[-" & i & "]"
        
    Next
    
    Ws.Cells(2, 15).Formula = "=RC[-1]/Kaikki!RC[-" & i & "]"
    Ws.Range("O3:O4").Merge
    Ws.Cells(3, 15).Formula = "=(RC[-1]+R[+1]C[-1])/(Kaikki!RC[-" & i & "]+Kaikki!R[+1]C[-" & i & "])"
    'Ws.Cells(4, 15).Formula = "=RC[-1]/Kaikki!RC[-" & i & "]"
        
    Ws.Cells(7, 15).Formula = "=RC[-1]/Kaikki!R[-5]C[-" & i & "]"
    Ws.Range("O8:O9").Merge
    Ws.Cells(8, 15).Formula = "=(RC[-1]+R[+1]C[-1])/(Kaikki!R[-5]C[-" & i & "]+Kaikki!R[-4]C[-" & i & "])"
    'Ws.Cells(9, 15).Formula = "=RC[-1]/Kaikki!R[-5]C[-" & i & "]"
       
    Ws.Range("C:C,E:E,G:G,I:I,K:K,M:M,O:O").NumberFormat = "0.00%"
    Ws.Range("C1,E1,G1,I1,K1,M1,O1") = "kaikista"
    Ws.Columns("A:A").EntireColumn.AutoFit
    
    Set CreateWorksheet = Ws
    
End Function



' Lasketaan kaikki esitykset ja kannatukset yhteensä
Function Kaikki() As Worksheet

    Dim ylioppilaskunnat As CPresenter
    Set ylioppilaskunnat = New CPresenter
    Call ylioppilaskunnat.Init("(", Sheets("MUUTOSESITYSEXCEL"))
    
    Sheets("Kaikki").Delete ' Comment this if there isn't such sheet
    Dim Ws As Worksheet
    Set Ws = Sheets.Add
    Ws.Name = "Kaikki"
    
    Ws.Cells(1, 1) = "Kaikki"
    Ws.Cells(2, 1) = "Esitykset"
    Ws.Cells(3, 1) = "Esityksiä läpi"
    Ws.Cells(4, 1) = "Esityksiä läpi muutoksin"
    Ws.Cells(5, 1) = "Läpäisy"
    Ws.Cells(1, 2) = "Lipa"
    Ws.Cells(1, 3) = "Tosu"
    Ws.Cells(1, 4) = "Talousarvio"
    Ws.Cells(1, 5) = "Kannanotot"
    Ws.Cells(1, 6) = "Yhteiskannanotto"
    Ws.Cells(1, 7) = "Ponnet"
    Ws.Cells(1, 8) = "Yht."
    Ws.Range("H2:H4").Formula = "=SUM(RC[-6]:RC[-1])"
    Ws.Range("B5:H5").Formula = "=SUM(R[-2]C:R[-1]C)/R[-3]C"
    Ws.Range("B5:H5").NumberFormat = "0.00%"
    
    Dim i As Integer
    Dim kohtienMaara As Integer
    kohtienMaara = ylioppilaskunnat.Kokouskohdat.Count
    For i = 1 To kohtienMaara
        Ws.Cells(2, i + 1) = ylioppilaskunnat.Kokouskohdat(i).esitykset
        Ws.Cells(3, i + 1) = ylioppilaskunnat.Kokouskohdat(i).esityksetLapi
        Ws.Cells(4, i + 1) = ylioppilaskunnat.Kokouskohdat(i).esityksetLapiMuutoksin
    Next

    Ws.Cells(1, 1) = "Kaikki yhteensä"
    Ws.Columns("A:A").EntireColumn.AutoFit
    
    Set Kaikki = Ws
    
End Function



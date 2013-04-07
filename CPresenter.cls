VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class for a presenter, usually a student union.

Private pNimi As String
Private pWorksheet As Worksheet
Private pKokouskohdat As Collection

' Regular getters and letters

Public Property Get Nimi() As String
    Nimi = pNimi
End Property
Public Property Let Nimi(Value As String)
    pNimi = Value
End Property
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property
Public Property Let Worksheet(Value As Worksheet)
    Set pWorksheet = Value
End Property
Public Property Get Kokouskohdat() As Collection
    Set Kokouskohdat = pKokouskohdat
End Property
Public Property Let Kokouskohdat(Value As Collection)
    Set pKokouskohdat = Value
End Property


' Initialization function. Don't use any of the letters!
' Nimi: Esittäjän nimi
' Worksheet: MUUTOSESITYSEXCEL
Public Sub Init(Nimi As String, Ws As Worksheet)
    Me.Nimi = Nimi
    Me.Worksheet = Ws
    Call CalculateAll
End Sub

' UNFINISHED ToString function
Public Sub ToString()
    MsgBox "Nimi: " & Me.Nimi
End Sub

' Calculate all the numbers from the spreadsheet
Private Sub CalculateAll()
    
    Me.Kokouskohdat = New Collection
    
    Me.Kokouskohdat.Add Calculate(10, 17) ' 1. Lipa
    Me.Kokouskohdat.Add Calculate(21, 56) ' 2. Tosu
    Me.Kokouskohdat.Add Calculate(63, 65) ' 3. Talousarvio
    Me.Kokouskohdat.Add Calculate(68, 79) ' 4. Kannanotot
    Me.Kokouskohdat.Add Calculate(83, 83) ' 5. Yhteiskannanotto
    Me.Kokouskohdat.Add Combine(Calculate(57, 59), Calculate(88, 91)) ' 6. Ponnet
    
    'MsgBox "Moi"
    
End Sub

' Calculate numbers for a single agenda item
Private Function Calculate(firstRow As Integer, lastRow As Integer) As CMotions
    
    Dim motions As CMotions
    Set motions = New CMotions
    
    Dim row As Integer
    For row = firstRow To lastRow
    
        Dim isEsitysFound, isKannatusFound As Integer
        isEsitysFound = InStr(Me.Worksheet.Cells(row, 2).Value, Me.Nimi)
        isKannatusFound = InStr(Me.Worksheet.Cells(row, 3).Value, Me.Nimi)
        If isEsitysFound <> 0 Then
            motions.esitykset = motions.esitykset + 1
            Dim isSuccess As Integer
            isSuccess = InStr(Me.Worksheet.Cells(row, 8).Value, "hyv")
            If isSuccess <> 0 Then
                motions.esityksetLapi = motions.esityksetLapi + 1
            Else
                isAmended = InStr(Me.Worksheet.Cells(row, 7).Value, "muut - ks. seur.")
                isSuccess = InStr(Me.Worksheet.Cells(row + 1, 8).Value, "hyv")
                If isAmended <> 0 And isSuccess <> 0 Then
                    motions.esityksetLapiMuutoksin = motions.esityksetLapiMuutoksin + 1
                End If
            End If
        End If
        If isKannatusFound <> 0 Then
            motions.kannatukset = motions.kannatukset + 1
            isSuccess = InStr(Me.Worksheet.Cells(row, 8).Value, "hyv")
            If isSuccess <> 0 Then
                motions.kannatuksetLapi = motions.kannatuksetLapi + 1
            Else
                isAmended = InStr(Me.Worksheet.Cells(row, 7).Value, "muut - ks. seur.")
                isSuccess = InStr(Me.Worksheet.Cells(row + 1, 8).Value, "hyv")
                If isAmended <> 0 And isSuccess <> 0 Then
                    motions.kannatuksetLapiMuutoksin = motions.kannatuksetLapiMuutoksin + 1
                End If
            End If
        End If
    
    Next row
    
    Set Calculate = motions
    
End Function

' Combine numbers of "two separate ranges" so to speak.
Private Function Combine(motions1 As CMotions, motions2 As CMotions) As CMotions
    Dim combination As CMotions
    Set combination = New CMotions
    combination.esitykset = motions1.esitykset + motions2.esitykset
    combination.esityksetLapi = motions1.esityksetLapi + motions2.esityksetLapi
    combination.esityksetLapiMuutoksin = motions1.esityksetLapiMuutoksin + motions2.esityksetLapiMuutoksin
    combination.kannatukset = motions1.kannatukset + motions2.kannatukset
    combination.kannatuksetLapi = motions1.kannatuksetLapi + motions2.kannatuksetLapi
    combination.kannatuksetLapiMuutoksin = motions1.kannatuksetLapiMuutoksin + motions2.kannatuksetLapiMuutoksin
    Set Combine = combination
End Function


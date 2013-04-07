VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMotions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Struct for the motions of an agenda item

Private pEsitykset As Integer
Private pKannatukset As Integer
Private pEsityksetLapi As Integer
Private pKannatuksetLapi As Integer
Private pEsityksetLapiMuutoksin As Integer
Private pKannatuksetLapiMuutoksin As Integer
Public Property Get esitykset() As Integer
    esitykset = pEsitykset
End Property
Public Property Let esitykset(Value As Integer)
    pEsitykset = Value
End Property
Public Property Get kannatukset() As Integer
    kannatukset = pKannatukset
End Property
Public Property Let kannatukset(Value As Integer)
    pKannatukset = Value
End Property
Public Property Get esityksetLapi() As Integer
    esityksetLapi = pEsityksetLapi
End Property
Public Property Let esityksetLapi(Value As Integer)
    pEsityksetLapi = Value
End Property
Public Property Get kannatuksetLapi() As Integer
    kannatuksetLapi = pKannatuksetLapi
End Property
Public Property Let kannatuksetLapi(Value As Integer)
    pKannatuksetLapi = Value
End Property
Public Property Get esityksetLapiMuutoksin() As Integer
    esityksetLapiMuutoksin = pEsityksetLapiMuutoksin
End Property
Public Property Let esityksetLapiMuutoksin(Value As Integer)
    pEsityksetLapiMuutoksin = Value
End Property
Public Property Get kannatuksetLapiMuutoksin() As Integer
    kannatuksetLapiMuutoksin = pKannatuksetLapiMuutoksin
End Property
Public Property Let kannatuksetLapiMuutoksin(Value As Integer)
    pKannatuksetLapiMuutoksin = Value
End Property

Attribute VB_Name = "Module1"
Public dbprenotazioni As Database
Public tabprenot As Recordset
Public prenotazioni As String


Public Const max = 200
Public OK As Boolean
Public i As Integer
Public codPrenotaz(1 To max) As String
Public nome(1 To max) As String
Public npartecipanti(1 To max) As Integer
Public terminal(1 To max) As String
Public provenienza(1 To max) As String
Public nvolo(1 To max) As String
Public tipotrasfer(1 To max) As String
Public trasferpriv(1 To max) As String
Public data(1 To max) As Date
Public ora(1 To max) As String
Public hotel(1 To max) As String
Public note(1 To max) As String


Public rigasel As Integer

Public finito As Boolean
Public ricercanome As String

Public caricato As Boolean



Public codiceritorno As String
Public tabRit As Recordset
Public dataritorno As Date
Public oraritorno As String

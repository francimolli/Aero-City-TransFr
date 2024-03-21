VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "SCHEDA DATE PRENOTAZIONI"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10260
   LinkTopic       =   "Form4"
   ScaleHeight     =   6195
   ScaleWidth      =   10260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "   ORDINA PER DATA        E PER ORA"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdStampa 
      Default         =   -1  'True
      Height          =   615
      Left            =   5520
      Picture         =   "DATE PRENOTAZ..frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "ESCI"
      Height          =   495
      Left            =   8520
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdClienti 
      Caption         =   "CLIENTI"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuovoCLI 
      Caption         =   "NUOVO CLIENTE"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdTorna 
      Caption         =   "TORNA"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia4 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type recOrdinamento
    ora As String
    data As Date
    codp As String
    nome As String
    part As Integer
    trasp As String
End Type
Dim ord(1 To max) As recOrdinamento
Dim temp As Variant


Private Sub cmdClienti_Click()
'VISUALIZZAZIONE
If tabprenot.RecordCount = 0 Then
    Me.Hide
    Form2.Show
Else
  If caricato = False Then
    Form2.Griglia.Rows = i + 1
    tabprenot.Index = "ixCodice"
    tabprenot.MoveFirst
    For j = 1 To i
        For k = 1 To i
            If IsNull(tabprenot!codPrenotaz) Then
            Else
                Form2.Griglia.TextMatrix(j, 0) = tabprenot!codPrenotaz
            End If
            If IsNull(tabprenot!nome) Then
            Else
                Form2.Griglia.TextMatrix(j, 1) = tabprenot!nome
            End If
            If IsNull(tabprenot!N°Partecipanti) Then
            Else
                Form2.Griglia.TextMatrix(j, 2) = tabprenot!N°Partecipanti
            End If
            If IsNull(tabprenot!terminal) Then
            Else
                Form2.Griglia.TextMatrix(j, 3) = tabprenot!terminal
            End If
            If IsNull(tabprenot!provenienza) Then
            Else
                Form2.Griglia.TextMatrix(j, 4) = tabprenot!provenienza
            End If
            If IsNull(tabprenot!N°volo) Then
            Else
                Form2.Griglia.TextMatrix(j, 5) = tabprenot!N°volo
            End If
            If IsNull(tabprenot!tipotrasfer) Then
            Else
                Form2.Griglia.TextMatrix(j, 6) = tabprenot!tipotrasfer
            End If
            If IsNull(tabprenot!data) Then
            Else
                Form2.Griglia.TextMatrix(j, 7) = tabprenot!data
            End If
            If IsNull(tabprenot!ora) Then
            Else
                Form2.Griglia.TextMatrix(j, 8) = Format(tabprenot!ora, "hh:mm")
            End If
            If IsNull(tabprenot!hotel) Then
            Else
                Form2.Griglia.TextMatrix(j, 9) = tabprenot!hotel
            End If
            If IsNull(tabprenot!note) Then
            Else
                Form2.Griglia.TextMatrix(j, 10) = tabprenot!note
            End If
        Next k
        tabprenot.MoveNext
    Next j
    caricato = True
    Me.Hide
    Form2.Show
  Else
    Me.Hide
    Form2.Show
  End If
End If
End Sub


Private Sub cmdEsci_Click()
    Unload Me
    End
End Sub

Private Sub cmdNuovoCLI_Click()
    Me.Hide
    Form1.Show
End Sub

Private Sub cmdStampa_Click()
    Me.PrintForm
End Sub

Private Sub cmdTorna_Click()
    Me.Hide
    Form3.Show
End Sub

Private Sub Command1_Click()
If tabprenot.RecordCount = 0 Then
    MsgBox "Nessun record presente"
Else
tabprenot.Index = "IXDataOra"
tabprenot.MoveFirst
j = 0
    Do While Not (tabprenot.EOF)
        j = j + 1
            If IsNull(tabprenot!ora) Then
            Else
                Griglia4.TextMatrix(j, 0) = Format(tabprenot!ora, "hh:mm")
            End If
            If IsNull(tabprenot!data) Then
            Else
                Griglia4.TextMatrix(j, 1) = tabprenot!data
            End If
            If IsNull(tabprenot!codPrenotaz) Then
            Else
                Griglia4.TextMatrix(j, 2) = tabprenot!codPrenotaz
            End If
            If IsNull(tabprenot!nome) Then
            Else
                Griglia4.TextMatrix(j, 3) = tabprenot!nome
            End If
            If IsNull(tabprenot!N°Partecipanti) Then
            Else
                Griglia4.TextMatrix(j, 4) = "1 + " & tabprenot!N°Partecipanti
            End If
            If IsNull(tabprenot!mezzotrasp) Then
            Else
                Griglia4.TextMatrix(j, 5) = tabprenot!mezzotrasp
            End If
            tabprenot.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()
    Griglia4.TextMatrix(0, 0) = "ORA"
    Griglia4.ColWidth(0) = 1500
    Griglia4.TextMatrix(0, 1) = "DATA"
    Griglia4.ColWidth(1) = 1500
    Griglia4.TextMatrix(0, 2) = "COD PRENOTAZ"
    Griglia4.ColWidth(2) = 1500
    Griglia4.TextMatrix(0, 3) = "NOME"
    Griglia4.ColWidth(3) = 2000
    Griglia4.TextMatrix(0, 4) = "N° PARTECIPANTI"
    Griglia4.ColWidth(4) = 1500
    Griglia4.TextMatrix(0, 5) = "TRASPORTO"
    Griglia4.ColWidth(5) = 1500
End Sub

'Private Sub cmdORDINAMENTO_Click()
'If tabprenot.RecordCount = 0 Then
'    MsgBox "Nessun record presente"
'Else
'tabprenot.MoveFirst
 '   For j = 1 To i
  '      ord(j).ora = tabprenot!ora
   '     ord(j).data = tabprenot!data
    '    ord(j).codp = tabprenot!codPrenotaz
     '   ord(j).nome = tabprenot!nome
      '  ord(j).part = tabprenot!N°Partecipanti
   '     ord(j).trasp = tabprenot!mezzotrasp
   '     tabprenot.MoveNext
   ' Next j
    
    'ORDINAMENTO X DATA CON BUBBLE SORT
    '    Do
      '      scambio = False
     '       For j = 1 To i - 1
       '         If ord(j).data > ord(j + 1).data Then
        '            temp = ord(j).ora
         '           ord(j).ora = ord(j + 1).ora
          '          ord(j + 1).ora = temp
           '
            '        temp = ord(j).data
             '       ord(j).data = ord(j + 1).data
              '      ord(j + 1).data = temp
               '
                '    temp = ord(j).codp
                 '   ord(j).codp = ord(j + 1).codp
                '    ord(j + 1).codp = temp
               '
              '      temp = ord(j).nome
             '       ord(j).nome = ord(j + 1).nome
            '        ord(j + 1).nome = temp
           '
          '          temp = ord(j).part
         '           ord(j).part = ord(j + 1).part
        '            ord(j + 1).part = temp
       '
      '              temp = ord(j).trasp
     '               ord(j).trasp = ord(j + 1).trasp
    '                ord(j + 1).trasp = temp
   '                 scambio = True
  '              End If
'            Next j
 '       Loop Until scambio = False
        
    'VISU
        'For j = 1 To i - 1
       '     Griglia4.TextMatrix(j, 0) = ord(j).ora
      '      Griglia4.TextMatrix(j, 1) = ord(j).data
     '       Griglia4.TextMatrix(j, 2) = ord(j).codp
    '        Griglia4.TextMatrix(j, 3) = ord(j).nome
   '         Griglia4.TextMatrix(j, 4) = ord(j).part
  '          Griglia4.TextMatrix(j, 5) = ord(j).trasp
 '       Next j
'End If
'End Sub

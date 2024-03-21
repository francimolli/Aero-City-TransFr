VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "VISUALIZZA DATA/ORA PARTENZA"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form7"
   ScaleHeight     =   5535
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOrdina 
      Caption         =   "ORDINA PER DATA/ORA"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ESCI"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TORNA"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia8 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOrdina_Click()
Set tabRit = dbprenotazioni.OpenRecordset("tabRitorno")
If tabRit.RecordCount = 0 Then
    MsgBox "Nessun record presente"
Else
tabRit.Index = "ixDataOra"
tabRit.MoveFirst
j = 0
    Do While Not (tabRit.EOF)
        j = j + 1
            If IsNull(tabRit!codicecliente) Then
            Else
                Griglia8.TextMatrix(j, 0) = tabRit!codicecliente
            End If
            If IsNull(tabRit!datacliente) Then
            Else
                Griglia8.TextMatrix(j, 1) = tabRit!datacliente
            End If
            If IsNull(tabRit!oracliente) Then
            Else
                Griglia8.TextMatrix(j, 2) = Format(tabRit!oracliente, "hh:mm")
            End If
            tabRit.MoveNext
    Loop
End If
End Sub

Private Sub Command1_Click()
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

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Griglia8.TextMatrix(0, 0) = "CODICE PRENOTAZIONE"
    Griglia8.ColWidth(0) = 2500
    Griglia8.TextMatrix(0, 1) = "DATA PARTENZA"
    Griglia8.ColWidth(1) = 2500
    Griglia8.TextMatrix(0, 2) = "ORA PARTENZA"
    Griglia8.ColWidth(2) = 2500
End Sub


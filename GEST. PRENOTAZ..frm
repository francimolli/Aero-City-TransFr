VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Gestione prenotazioni trasporto da aeroporto:"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3360
      TabIndex        =   32
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAzzera 
      Caption         =   "Azzera input"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton cmdTrasporto 
      Caption         =   "Trasporto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   15
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisualizza 
      Caption         =   "Visualizza clienti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "Conferma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuovo cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtHotel 
         Height          =   435
         Left            =   4320
         TabIndex        =   9
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox txtNote 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   6120
         Width           =   5655
      End
      Begin VB.ComboBox cmbPrivato 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "GEST. PRENOTAZ..frx":0000
         Left            =   5280
         List            =   "GEST. PRENOTAZ..frx":000D
         TabIndex        =   29
         Text            =   "Limusine (1-3)"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtCodPrenotaz 
         Height          =   435
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbTipoTrasfer 
         Height          =   315
         ItemData        =   "GEST. PRENOTAZ..frx":003C
         Left            =   3600
         List            =   "GEST. PRENOTAZ..frx":0046
         TabIndex        =   6
         Text            =   "Condiviso"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtNVolo 
         Height          =   435
         Left            =   4320
         TabIndex        =   5
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtProvenienza 
         Height          =   435
         Left            =   4320
         TabIndex        =   4
         Top             =   2640
         Width           =   2775
      End
      Begin VB.ComboBox cmbTerminal 
         Height          =   315
         ItemData        =   "GEST. PRENOTAZ..frx":005E
         Left            =   4320
         List            =   "GEST. PRENOTAZ..frx":006E
         TabIndex        =   3
         Text            =   "Malpensa T1"
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtOra 
         Height          =   435
         Left            =   4320
         TabIndex        =   8
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox txtData 
         Height          =   435
         Left            =   4320
         TabIndex        =   7
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox txtNPartecipanti 
         Height          =   435
         Left            =   4320
         TabIndex        =   2
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtNome 
         Height          =   435
         Left            =   4320
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label14 
         Caption         =   "HOTEL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "NOTE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "COD PRENOTAZ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "TIPO DI TRASFER."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "N° VOLO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "TERMINAL PROVENIENZA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "(HH.mm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "ORA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "DATA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "(dd/mm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "TERMINAL ARRIVO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "N°  ALTRI  PARTECIPANTI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "NOME:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANC As Boolean

Private Sub cmbTipoTrasfer_Click()
Select Case cmbTipoTrasfer.ListIndex
        Case 0
            cmbPrivato.Enabled = False
        Case 1
            cmbPrivato.Enabled = True
                If i = 0 Then
                    i = 1
                End If
            Select Case cmbPrivato.ListIndex
                Case -1
                    trasferpriv(i) = cmbPrivato.Text
                Case 0
                    trasferpriv(i) = "Limusine (1-3)"
                Case 1
                    trasferpriv(i) = "Mini BUS (6-8)"
                Case 2
                    trasferpriv(i) = "BUS (~50)"
            End Select
                If i = 1 Then
                    i = 0
                End If
End Select
End Sub




Private Sub Form_Load()
    i = 0
    prenotazioni = App.Path & "\GestionePrenotazioni.mdb"
    Set dbprenotazioni = OpenDatabase(prenotazioni)
    Set tabprenot = dbprenotazioni.OpenRecordset("tabprenotaz")
    Do While Not (tabprenot.EOF)
        tabprenot.MoveNext
        i = i + 1
    Loop
    txtCodPrenotaz.Text = "C" & (i + 1)
End Sub

Private Sub cmdConferma_Click()
CANC = True
caricato = False
    
    'CONTATORE ARRAY
    i = i + 1
    
    'CARICAMENTO DB

    
    ' COD PRENOTAZIONE
If tabprenot.RecordCount = 0 Then
        tabprenot.AddNew
                codPrenotaz(i) = txtCodPrenotaz.Text
                tabprenot("codprenotaz") = codPrenotaz(i)
        OK = True
Else
    If txtCodPrenotaz.Text <> "" Then
        OK = True
        codPrenotaz(i) = txtCodPrenotaz.Text
        tabprenot.MoveFirst
        For j = 1 To i - 1
            If codPrenotaz(i) = tabprenot!codPrenotaz Then
                MsgBox "IMPOSSIBILE DUPLICARE INDICE CHIAVE PRIMARIA", vbCritical, "ERRORE"
                OK = False
                CANC = False
            End If
            tabprenot.MoveNext
        Next j
If OK = True Then
        tabprenot.AddNew
        tabprenot("codprenotaz") = codPrenotaz(i)
End If
    Else
        MsgBox "inserire codice prenotazione valido", , ""
    End If
End If
    'DATA
    
If OK = True Then
    If txtData.Text <> "" Then
        data(i) = txtData.Text
        tabprenot("data") = data(i)
    Else
        MsgBox "data non inserita!", , ""
        OK = False
        CANC = False
    End If
End If
    
    'ORA
If OK = True Then
    If txtOra.Text <> "" Then
        ora(i) = txtOra.Text
        tabprenot("ora") = ora(i)
    Else
        MsgBox "ora di arrivo non inserita!", , ""
        OK = False
        CANC = False
    End If
End If
    
    'NOME CLIENTE
    
If OK = True Then
    If txtNome.Text <> "" Then
        nome(i) = txtNome.Text
        tabprenot("nome") = nome(i)
    Else
        MsgBox "nome non inserito!", , ""
    End If
End If

    'NUMERO ALTRI PARTECIPANTI
    
If OK = True Then
    If txtNPartecipanti.Text <> "" Then
        npartecipanti(i) = txtNPartecipanti.Text
        tabprenot("N°Partecipanti") = npartecipanti(i)
    End If
End If

    'TERMINAL ARRIVO
    
If OK = True Then
    Select Case cmbTerminal.ListIndex
        Case -1
            terminal(i) = cmbTerminal.Text
        Case 0
            terminal(i) = "Malpensa T1"
        Case 1
            terminal(i) = "Malpensa T2"
        Case 2
            terminal(i) = "Linate"
        Case 3
            terminal(i) = "Orio al Serio"
        Case Else
            MsgBox "ERRORE", vbCritical, ""
    End Select
    tabprenot("terminal") = terminal(i)
End If
    
    'TERMINAL PROVENIENZA

If OK = True Then
    If txtProvenienza.Text <> "" Then
        provenienza(i) = txtProvenienza.Text
        tabprenot("provenienza") = provenienza(i)
    Else
        MsgBox "luogo di provenienza valido non inserito ", , ""
    End If
End If
    
    'NUMERO VOLO
    
If OK = True Then
    If txtNVolo.Text <> "" Then
        nvolo(i) = txtNVolo.Text
        tabprenot("N°Volo") = nvolo(i)
    Else
        MsgBox "numero di volo valido non inserito", , ""
    End If
End If
    
    ' TIPO DI TRASFERTA
    

    Select Case cmbTipoTrasfer.ListIndex
        Case -1
            tipotrasfer(i) = cmbTipoTrasfer.Text
        Case 0
            tipotrasfer(i) = "Condiviso"
        Case 1
            tipotrasfer(i) = "Privato"
            cmbPrivato.Enabled = True
            
            Select Case cmbPrivato.ListIndex
                Case -1
                    trasferpriv(i) = cmbPrivato.Text
                Case 0
                    trasferpriv(i) = "Limusine (1-3)"
                Case 1
                    trasferpriv(i) = "Mini BUS (6-8)"
                Case 2
                    trasferpriv(i) = "BUS (~50)"
            End Select
        Case Else
            MsgBox "ERRORE", vbCritical, ""
    End Select
If OK = True Then
    tabprenot("tipotrasfer") = tipotrasfer(i)
    tabprenot("MezzoTrasp") = trasferpriv(i)
End If

    'HOTEL
If OK = True Then
    If txtHotel <> "" Then
        hotel(i) = txtHotel.Text
        tabprenot("hotel") = hotel(i)
    End If
End If

    'NOTE
If OK = True Then
    If txtNote <> "" Then
        note(i) = txtNote.Text
        tabprenot("note") = note(i)
    End If
End If
    
    'CARICAMENTO DB E AGGIORNAMENTO ROW GRIGLIA X FORM2 DI VISUALIZZAZIONE
    
If OK = True Then
        tabprenot.Update
    Else
        i = i - 1
End If
If CANC = True Then
txtCodPrenotaz.Text = "C" & (i + 1)
txtNome.Text = ""
txtNPartecipanti.Text = ""
txtProvenienza.Text = ""
txtNVolo.Text = ""
txtData.Text = ""
txtOra.Text = ""
txtHotel.Text = ""
txtNote.Text = ""
End If
End Sub

Private Sub cmdEsci_Click()
    Unload Me
    End
End Sub

Private Sub cmdVisualizza_Click()
'VISUALIZZAZIONE
If tabprenot.RecordCount = 0 Then
    Me.Hide
    Form2.Show
Else
  If caricato = False Then
    MsgBox "è in corso il caricamento dei dati dal Database, l'operazione potrà richiedere alcuni istanti", , ""
    cmdTrasporto.Visible = False
    ProgressBar1.Visible = True
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
        ProgressBar1.Value = (100 / i) * j
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
Private Sub cmdTrasporto_Click()
If tabprenot.RecordCount = 0 Then
    Me.Hide
    Form3.Show
    Form3.Griglia2.TextMatrix(1, 0) = ""
    Form3.Griglia2.TextMatrix(1, 1) = ""
    Form3.Griglia2.TextMatrix(1, 2) = ""
    Form3.Griglia2.TextMatrix(1, 3) = ""
    Form3.Griglia2.TextMatrix(1, 4) = ""
Else
    Form3.Griglia2.Rows = i + 1
    Me.Hide
    Form3.Show
    tabprenot.Index = "ixCodice"
    tabprenot.MoveFirst
    For j = 1 To i
        For k = 1 To i
            If IsNull(tabprenot!codPrenotaz) Then
            Else
                Form3.Griglia2.TextMatrix(j, 0) = tabprenot!codPrenotaz
            End If
            If IsNull(tabprenot!nome) Then
            Else
                Form3.Griglia2.TextMatrix(j, 1) = tabprenot!nome
            End If
            If IsNull(tabprenot!N°Partecipanti) Then
            Else
                Form3.Griglia2.TextMatrix(j, 2) = "1 + " & tabprenot!N°Partecipanti
            End If
            If IsNull(tabprenot!tipotrasfer) Then
            Else
                Form3.Griglia2.TextMatrix(j, 3) = tabprenot!tipotrasfer
            End If
            If IsNull(tabprenot!mezzotrasp) Then
            Else
                Form3.Griglia2.TextMatrix(j, 4) = tabprenot!mezzotrasp
            End If
        Next k
        tabprenot.MoveNext
    Next j
End If
End Sub

Private Sub cmdAzzera_Click()
    txtCodPrenotaz.Text = ""
    txtNome.Text = ""
    txtNPartecipanti.Text = ""
    txtProvenienza.Text = ""
    txtNVolo.Text = ""
    txtData.Text = ""
    txtOra.Text = ""
    txtHotel.Text = ""
    txtNote.Text = ""
End Sub


VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Visualizza prenotazioni"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14310
   LinkTopic       =   "Form2"
   ScaleHeight     =   6975
   ScaleWidth      =   14310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRitorno 
      Caption         =   "IMPOSTA DATA/ORA RITORNO"
      Height          =   615
      Left            =   10560
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdRicerca1 
      Caption         =   "RICERCA PER NOME"
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdModifica 
      Caption         =   "MODIFICA DATI CLIENTE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdStampa 
      Default         =   -1  'True
      Height          =   615
      Left            =   7440
      Picture         =   "VISU. PRENOTAZ..frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   615
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "ELIMINA CLIENTE"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "ESCI"
      Height          =   495
      Left            =   12000
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdNuovoCliente 
      Caption         =   "NUOVO CLIENTE"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   6240
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim risposta As String


Private Sub cmdElimina_Click()
    Dim strCodPren As String
    ' PROCEDURA ELIMINAZIONE RECORD/CLIENTE
    If tabprenot.RecordCount = 0 Then
        MsgBox "Nessun record corrente!"
    Else
    If Griglia.RowSel > 0 Then
        risposta = MsgBox("Sei sicuro di voler cancellare definitivamente il cliente selezionato?", vbQuestion + vbYesNo, "Elimina")
        If risposta = vbYes Then
            Set tabprenot = dbprenotazioni.OpenRecordset("tabprenotaz")
            tabprenot.MoveFirst
            
            ' IL TUO ERRORE (IL + 1 VA TOLTO) **********
            'tabprenot.Move (rigasel + 1)
            tabprenot.Move (rigasel)
            ' ******************************************
            
            ' Meglio però fare così:
            rigasel = Griglia.RowSel
            strCodPren = Griglia.TextMatrix(rigasel, 0)
            tabprenot.Index = "PrimaryKey"
            tabprenot.Seek "=", strCodPren
            ' ******************************************
            
            MsgBox "Cancello:  " & tabprenot!codPrenotaz
            MsgBox "attendere, cancellazione in corso"
            tabprenot.Delete
            If Griglia.Row = 1 Then
                Griglia.TextMatrix(1, 0) = ""
                Griglia.TextMatrix(1, 1) = ""
                Griglia.TextMatrix(1, 2) = ""
                Griglia.TextMatrix(1, 3) = ""
                Griglia.TextMatrix(1, 4) = ""
                Griglia.TextMatrix(1, 5) = ""
                Griglia.TextMatrix(1, 6) = ""
                Griglia.TextMatrix(1, 7) = ""
                Griglia.TextMatrix(1, 8) = ""
            Else
                Griglia.RemoveItem (Griglia.RowSel)
            End If
            i = i - 1
        Else
            MsgBox "Calcellazione annullata"
        End If
    Else: MsgBox "Selezione il cliente che si desidera eliminare"
    End If
    End If
End Sub

Private Sub cmdRicerca1_Click()
    ricercanome = InputBox("", "Inserisci nome da ricercare")
    Set tabprenot = dbprenotazioni.OpenRecordset("tabprenotaz")
    tabprenot.Index = "ixNome"
    tabprenot.Seek "=", ricercanome
    If tabprenot.NoMatch Then
        MsgBox "VALORE NON TROVATO!", , ""
    Else
        Form5.Griglia6.TextMatrix(1, 0) = tabprenot!codPrenotaz
        Form5.Griglia6.TextMatrix(1, 1) = tabprenot!nome
        Form5.Show
    End If
End Sub

Private Sub cmdRitorno_Click()
    codiceritorno = InputBox("", "Inserire codice cliente")
    Set tabprenot = dbprenotazioni.OpenRecordset("tabprenotaz")
    tabprenot.Index = "ixCodice"
    tabprenot.Seek "=", codiceritorno
    If tabprenot.NoMatch Then
        MsgBox "CODICE NON TROVATO!", , ""
    Else
        Form6.lblCodice.Caption = tabprenot!codPrenotaz
        Form6.lblNome.Caption = tabprenot!nome
        Form6.Show
    End If
End Sub

'Private Sub cmdModifica_Click()
'If tabprenot.RecordCount = 0 Then
 '       MsgBox "Nessun record corrente!"
  '  Else
   ' If Griglia.RowSel > 0 Then
    '    risposta = MsgBox("Sei sicuro di voler modificare i dati del cliente selezionato?", vbQuestion + vbYesNo, "Elimina")
     '   If risposta = vbYes Then
      '      Set tabprenot = dbprenotazioni.OpenRecordset("tabprenotaz")
       '     tabprenot.MoveFirst
        '    tabprenot.Move (rigasel - 1)
         '   MsgBox "Modifica di:  " & tabprenot!codPrenotaz
          '  MsgBox "attendere, procedura di modifica in corso"
           ' tabprenot.Edit
            'Form1.Show
            'F orm1.cmdAzzera.Visible = False
           ' Form1.cmdConferma.Visible = False
           ' Form1.cmdEsci.Visible = False
           ' Form1.cmdTrasporto.Visible = False
'            Form1.cmdVisualizza.Visible = False
 '           Form1.txtCodPrenotaz.Text = Griglia.TextMatrix(rigasel, 0)
  '          Form1.txtNome.Text = Griglia.TextMatrix(rigasel, 1)
   '         Form1.txtNPartecipanti.Text = Griglia.TextMatrix(rigasel, 2)
    '        Form1.cmbTerminal.Text = Griglia.TextMatrix(rigasel, 3)
     '       Form1.txtProvenienza.Text = Griglia.TextMatrix(rigasel, 4)
      '      Form1.txtNVolo.Text = Griglia.TextMatrix(rigasel, 5)
       '     Form1.cmbTipoTrasfer.Text = Griglia.TextMatrix(rigasel, 6)
        '    Form1.txtData.Text = Griglia.TextMatrix(rigasel, 7)
         '   Form1.txtOra.Text = Griglia.TextMatrix(rigasel, 8)
          '  Form1.txtNote.Text = Griglia.TextMatrix(rigasel, 9)
           ' Form1.cmdApporta.Visible = True
            'Form1.cmdAnnulla.Visible = True
'        Else
 '           MsgBox "Modifica annullata"
  '      End If
   ' Else: MsgBox "Selezione il cliente cui si desidera modificare i dati"
    'End If
  '  End If
'End Sub

Private Sub cmdStampa_Click()
    Me.PrintForm
End Sub

Private Sub Griglia_Click()
    If Griglia.RowSel > 0 Then
        rigasel = Griglia.RowSel
    Else
        MsgBox "Seleziona una riga!"
    End If
        
End Sub

Private Sub cmdEsci_Click()
    Unload Me
    End
End Sub

Private Sub cmdNuovoCliente_Click()
    Form1.ProgressBar1.Visible = False
    Form1.cmdTrasporto.Visible = True
    Me.Hide
    Form1.Show
End Sub


Private Sub Form_Load()
    'INTESTA GRIGLIA
    Griglia.TextMatrix(0, 0) = "COD PRENOTAZ"
    Griglia.ColWidth(0) = 1500
    Griglia.TextMatrix(0, 1) = "NOME"
    Griglia.ColWidth(1) = 2500
    Griglia.TextMatrix(0, 2) = "N° ALTRI PARTEC."
    Griglia.ColWidth(2) = 1700
    Griglia.TextMatrix(0, 3) = "TERMINAL ARRIVO"
    Griglia.ColWidth(3) = 1700
    Griglia.TextMatrix(0, 4) = "TERMINAL PROVENIENZA"
    Griglia.ColWidth(4) = 3500
    Griglia.TextMatrix(0, 5) = "N° VOLO"
    Griglia.ColWidth(5) = 2500
    Griglia.TextMatrix(0, 6) = "TIPO TRANSF"
    Griglia.ColWidth(6) = 1400
    Griglia.TextMatrix(0, 7) = "DATA"
    Griglia.TextMatrix(0, 8) = "ORA (hh.mm.ss)"
    Griglia.ColWidth(8) = 1400
    Griglia.TextMatrix(0, 9) = "HOTEL"
    Griglia.ColWidth(9) = 3500
    Griglia.TextMatrix(0, 10) = "NOTE"
    Griglia.ColWidth(10) = 8000
End Sub

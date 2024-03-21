VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "VISU. TIPO TRASPORTO"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form3"
   ScaleHeight     =   6435
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStampa 
      Default         =   -1  'True
      Height          =   615
      Left            =   6960
      Picture         =   "VISU. TRASPORTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdSchedaDATE 
      Caption         =   "SCHEDA DATE"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton cmdClienti 
      Caption         =   "CLIENTI"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsci 
      Caption         =   "ESCI"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdNCli 
      Caption         =   "NUOVO CLIENTE"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia2 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdNCli_Click()
    Me.Hide
    Form1.Show
End Sub

Private Sub cmdSchedaDATE_Click()
If tabprenot.RecordCount = 0 Then
    Me.Hide
    Form4.Show
    Form4.Griglia4.TextMatrix(1, 0) = ""
    Form4.Griglia4.TextMatrix(1, 1) = ""
    Form4.Griglia4.TextMatrix(1, 2) = ""
    Form4.Griglia4.TextMatrix(1, 3) = ""
    Form4.Griglia4.TextMatrix(1, 4) = ""
    Form4.Griglia4.TextMatrix(1, 5) = ""
Else
    Form4.Griglia4.Rows = i + 1
    Me.Hide
    Form4.Show
    tabprenot.MoveFirst
    For j = 1 To i
        For k = 1 To i
        
            If IsNull(tabprenot!ora) Then
            Else
                Form4.Griglia4.TextMatrix(j, 0) = Format(tabprenot!ora, "hh:mm")
            End If
            
            If IsNull(tabprenot!data) Then
            Else
                Form4.Griglia4.TextMatrix(j, 1) = tabprenot!data
            End If

            If IsNull(tabprenot!codPrenotaz) Then
            Else
                Form4.Griglia4.TextMatrix(j, 2) = tabprenot!codPrenotaz
            End If
            
            If IsNull(tabprenot!nome) Then
            Else
                Form4.Griglia4.TextMatrix(j, 3) = tabprenot!nome
            End If
            
            If IsNull(tabprenot!N°Partecipanti) Then
            Else
                Form4.Griglia4.TextMatrix(j, 4) = "1 + " & tabprenot!N°Partecipanti
            End If
            
            If IsNull(tabprenot!mezzotrasp) Then
            Else
                Form4.Griglia4.TextMatrix(j, 5) = tabprenot!mezzotrasp
            End If
            
        Next k
        tabprenot.MoveNext
    Next j
End If
End Sub

Private Sub cmdStampa_Click()
    Me.PrintForm
End Sub

Private Sub Form_Load()
    Griglia2.TextMatrix(0, 0) = "COD PRENOTAZ"
    Griglia2.ColWidth(0) = 1500
    Griglia2.TextMatrix(0, 1) = "NOME"
    Griglia2.ColWidth(1) = 2500
    Griglia2.TextMatrix(0, 2) = "N° PARTECIPANTI"
    Griglia2.ColWidth(2) = 1500
    Griglia2.TextMatrix(0, 3) = "TRASPORTO"
    Griglia2.ColWidth(3) = 1500
    Griglia2.TextMatrix(0, 4) = "TIPO TRASP."
    Griglia2.ColWidth(4) = 1500
End Sub

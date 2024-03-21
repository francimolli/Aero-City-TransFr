VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Inserimento Data di Ritorno"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form6"
   ScaleHeight     =   3480
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVisuRitorno 
      Caption         =   "VISUALIZZA DATE DI RITORNO"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "CHIUDI"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdComferma 
      Caption         =   "CONFERMA"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtOraRit 
         Height          =   435
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtDataRit 
         Height          =   435
         Left            =   2760
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label4 
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
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblNome 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblCodice 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label9 
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
         Left            =   1080
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   2160
         Width           =   1335
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
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChiudi_Click()
    Me.Hide
End Sub

Private Sub cmdComferma_Click()
    Set tabRit = dbprenotazioni.OpenRecordset("tabRitorno")
    dataritorno = txtDataRit.Text
    oraritorno = txtOraRit.Text
    tabRit.MoveLast
    tabRit.AddNew
    tabRit!codicecliente = tabprenot!codPrenotaz
    tabRit!datacliente = dataritorno
    tabRit!oracliente = oraritorno
    MsgBox "Inserimento completato!", , ""
    txtDataRit.Text = ""
    txtOraRit.Text = ""
    tabRit.Update
    Me.Hide
End Sub

Private Sub cmdVisuRitorno_Click()
Set tabRit = dbprenotazioni.OpenRecordset("tabRitorno")
 If tabRit.RecordCount = 0 Then
    Me.Hide
    Form7.Show
Else
    Form7.Griglia8.Rows = tabRit.RecordCount + 1
    Me.Hide
    Form7.Show
    tabRit.Index = "ixCodice"
    tabRit.MoveFirst
    For j = 1 To tabRit.RecordCount
        For k = 1 To tabRit.RecordCount
            If IsNull(tabRit!codicecliente) Then
            Else
                Form7.Griglia8.TextMatrix(j, 0) = tabRit!codicecliente
            End If
            If IsNull(tabRit!datacliente) Then
            Else
                Form7.Griglia8.TextMatrix(j, 1) = tabRit!datacliente
            End If
            If IsNull(tabRit!oracliente) Then
            Else
                Form7.Griglia8.TextMatrix(j, 2) = Format(tabRit!oracliente, "hh:mm")
            End If
        Next k
        tabRit.MoveNext
    Next j
End If
End Sub


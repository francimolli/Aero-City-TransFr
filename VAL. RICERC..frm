VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "VALORI RICERCATI"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form5"
   ScaleHeight     =   1305
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "CHIUDI"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid Griglia6 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChiudi_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Griglia6.TextMatrix(0, 0) = "CODICE PRENOTAZIONE"
    Griglia6.ColWidth(0) = 2500
    Griglia6.TextMatrix(0, 1) = "NOME"
    Griglia6.ColWidth(1) = 2500
End Sub

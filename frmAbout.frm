VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre..."
   ClientHeight    =   6525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7365
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4503.672
   ScaleMode       =   0  'User
   ScaleWidth      =   6916.117
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":0442
      Top             =   3120
      Width           =   6855
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5280
      TabIndex        =   0
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Line Line1 
      X1              =   225.372
      X2              =   6648.487
      Y1              =   1987.828
      Y2              =   1987.828
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   4560
      Picture         =   "frmAbout.frx":0616
      Top             =   120
      Width           =   2550
   End
   Begin VB.Label lblTitle 
      Caption         =   "Ricardo Lobo lista generator Tabajara"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   330
      TabIndex        =   2
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versão 0.000000000001"
      Height          =   225
      Left            =   330
      TabIndex        =   3
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":7A1D
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   375
      TabIndex        =   1
      Top             =   1305
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblTitle.Caption = App.Title
End Sub

VERSION 5.00
Begin VB.Form frmProductKey 
   BorderStyle     =   1  '단일 고정
   Caption         =   "DecoSteel - Product Key Check"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmProductKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confirm"
      Height          =   435
      Left            =   3930
      TabIndex        =   3
      Top             =   990
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   6015
      Begin VB.TextBox txtSerial 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   4155
      End
      Begin VB.TextBox txtPK 
         Height          =   270
         Left            =   1710
         TabIndex        =   2
         Top             =   600
         Width           =   4185
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   90
         Top             =   180
         Width           =   5805
      End
      Begin VB.Label label3 
         AutoSize        =   -1  'True
         Caption         =   "Serial Key        : "
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Certification Key:"
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   660
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmProductKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
gs_CertiKey = CStr(Trim(frmProductKey.txtPK.Text))
gin_Int = 1
Unload Me
End Sub

Private Sub Form_Load()
txtSerial.Text = gs_PW

End Sub

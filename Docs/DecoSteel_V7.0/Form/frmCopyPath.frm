VERSION 5.00
Begin VB.Form frmCopyPath 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Copy Path Check Window"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "frmCopyPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   8160
   Begin VB.Frame Frame2 
      Caption         =   "MicroStation Version Check"
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   5595
      Begin VB.OptionButton Option2 
         Caption         =   "MicroStation SE Version"
         Height          =   315
         Left            =   2850
         TabIndex        =   5
         Top             =   240
         Width           =   2505
      End
      Begin VB.OptionButton Option1 
         Caption         =   "MicroStation J Version"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Copy"
      Height          =   435
      Left            =   5760
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy Path"
      Height          =   765
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   8145
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   300
         Width           =   7845
      End
   End
End
Attribute VB_Name = "frmCopyPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstr_SourcePath As String

Private Sub cmdOK_Click()
Dim object
Dim FileName As String
On Error GoTo labelError


Set object = CreateObject("Scripting.FileSystemObject")

If gin_MenuFlag = 1 Then
            FileName = "Deco_ustn.m01"
 Else
            FileName = "ustn.m01"
 End If

If gin_MenuFlag = 1 Then
            Kill CStr(Trim(txtPath.Text)) & "\ustn.m01"
End If

object.CopyFile lstr_SourcePath, CStr(Trim(txtPath.Text)) & "\" & FileName, True

Set object = Nothing


If gin_MenuFlag = 1 Then
            Name CStr(Trim(txtPath.Text)) & "\Deco_ustn.m01" As CStr(Trim(txtPath.Text)) & "\ustn.m01"
End If

MsgBox "File Copy Success !!!"

If Len(Command) = 0 Then
            Unload Me
Else
            End
End If

Exit Sub

labelError:
MsgBox "File Copy Erro !!!"
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If

End Sub

Private Sub Form_Load()


Select Case gin_MenuFlag
            Case 1
                        lstr_SourcePath = App.Path & "\Menu_Deco\j_version\Deco_ustn.m01"
                        txtPath.Text = "C:\Bentley\Workspace\interfaces\MicroStation\default"
                        Me.Caption = "DecoSteel Menu Copy....."
            Case 2
                         lstr_SourcePath = App.Path & "\Menu_Ustation\j_version\ustn.m01"
                        txtPath.Text = "C:\Bentley\Workspace\interfaces\MicroStation\default"
                        Me.Caption = "MicroStation Menu Copy....."
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If

End Sub


Private Sub Option1_Click()

txtPath.Text = "C:\Bentley\Workspace\interfaces\MicroStation\default"
Select Case gin_MenuFlag
            Case 1
                        lstr_SourcePath = App.Path & "\Menu_Deco\j_version\Deco_ustn.m01"
            Case 2
                         lstr_SourcePath = App.Path & "\Menu_Ustation\j_version\ustn.m01"
End Select

End Sub

Private Sub Option2_Click()
txtPath.Text = "Not Support yet"
Select Case gin_MenuFlag
            Case 1
                        lstr_SourcePath = App.Path & "\Menu_Deco\se_version\Deco_ustn.m01"
            Case 2
                         lstr_SourcePath = App.Path & "\Menu_Ustation\se_version\ustn.m01"
End Select
End Sub

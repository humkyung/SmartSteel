VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBasePlate_Hinged 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Base Plate(Hinged Type)"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "FrmBasePlate_Hinged.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ImageList ImageList 
      Left            =   60
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   47
      ImageHeight     =   46
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBasePlate_Hinged.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBasePlate_Hinged.frx":1A3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   945
      TabIndex        =   13
      Top             =   4140
      Width           =   3300
      Begin VB.TextBox txtDM 
         Height          =   270
         Left            =   1890
         TabIndex        =   17
         Text            =   "A"
         Top             =   135
         Width           =   1365
      End
      Begin VB.TextBox txtRDN 
         Height          =   270
         Left            =   1890
         TabIndex        =   16
         Text            =   "AC-15"
         Top             =   405
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Reference DWG No :"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   450
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Detail Mark            :"
         Height          =   180
         Left            =   90
         TabIndex        =   14
         Top             =   225
         Width           =   1725
      End
   End
   Begin VB.CheckBox chkNut 
      Caption         =   "Nut Check (One Nut)"
      Height          =   420
      Left            =   45
      TabIndex        =   12
      Top             =   2115
      Value           =   1  '확인
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select PML Unit"
      Height          =   780
      Left            =   0
      TabIndex        =   7
      Top             =   45
      Width           =   3345
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   2610
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   1710
         TabIndex        =   10
         Top             =   360
         Width           =   690
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   1035
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   225
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5670
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make PML"
      Height          =   405
      Left            =   3465
      TabIndex        =   5
      Top             =   2640
      Width           =   2130
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Shape Option"
      Height          =   1185
      Left            =   90
      TabIndex        =   2
      Top             =   855
      Width           =   3300
      Begin VB.OptionButton Opt02 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1845
         TabIndex        =   4
         Top             =   360
         Width           =   240
      End
      Begin VB.OptionButton Opt01 
         Caption         =   "Option1"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   285
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   2205
         Top             =   360
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   585
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Column"
      Height          =   2490
      Left            =   3420
      TabIndex        =   0
      Top             =   45
      Width           =   3345
      Begin VB.ListBox lstColumn 
         Columns         =   2
         Height          =   2040
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   3165
      End
   End
End
Attribute VB_Name = "FrmBasePlate_Hinged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstr_Job As String, gs_PMLunit As String
Private Sub cmdExit_Click()
Dim i As Integer
'For i = CInt(Trim(Me.Width)) To 0 Step -1
'               Me.Width = i
'               Me.Height = i
'Next i

If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If

End Sub

Private Sub cmdMake_Click()

Dim TempPath As String, MemberName As String
Dim TempUnit As String, TempDir As String, TempNut As Integer
Dim TempDM As String, TempRDN As String

On Error GoTo Labelstop

If optMM.Value = True Then
    TempUnit = "mm"
ElseIf optM.Value = True Then
    TempUnit = "m"
ElseIf optInch.Value = True Then
    TempUnit = "inch"
ElseIf optFeet.Value = True Then
    TempUnit = "feet"
End If

If opt01.Value = True Then
    TempDir = "VectorY"
Else
    TempDir = "VectorX"
End If
TempNut = CInt(chkNut.Value)

MemberName = Trim(lstColumn.List(lstColumn.ListIndex))

TempDM = CStr(Trim(txtDM.Text))
TempRDN = CStr(Trim(txtRDN.Text))

If MemberName = "" Then
    MsgBox "Select Column Member Size....."
Else

    frmMain.CommonDialog.CancelError = True
    frmMain.CommonDialog.InitDir = App.Path
    frmMain.CommonDialog.DialogTitle = "Save PML File "
    frmMain.CommonDialog.Filter = "BasePlate (*.pml)|*.pml|"
    frmMain.CommonDialog.FileName = "Test.pml"
    
    frmMain.CommonDialog.ShowSave
    
    TempPath = frmMain.CommonDialog.FileName
    Call BasePlate_Hinged_PML(TempPath, gstr_Job, MemberName, TempUnit, TempDir, TempNut, TempDM, TempRDN)
    
    Call PML_Run(TempPath)
    End
End If
Labelstop:
    
End Sub

Private Sub Form_Load()
Dim sql As String
Dim i

Set Image1.Picture = ImageList.ListImages(1).Picture
Set Image2.Picture = ImageList.ListImages(2).Picture

If gin_Chk_Flag01 = 0 Then
             i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
            i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

If Len(Command) <> 0 Then
      Dim lstr_ProjectName As String, lin_Error As Integer
      Call gs_Call_Project(lstr_ProjectName, lin_Error)
      If lin_Error = 1 Then MsgBox "Job is not selected. Please Job Select at General": End
      gstr_Job = lstr_ProjectName
      Me.Caption = "Current Project : " & gstr_Job & " , Base Plate(Hinged Type)"
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
            
End If

If gstr_Job = "" Then
      MsgBox "Job is not selected. Please Job Select at General"
      End
Else

      sql = "select member_name from BasePlate_Hinged where job = '" & gstr_Job & "'"
      Call Query_AddList_function(1, lstColumn, sql)
End If

If gs_PMLunit = "mm" Then
      optMM.Value = True
      optM.Value = False
      optInch.Value = False
      optFeet.Value = False
ElseIf gs_PMLunit = "m" Then
      optMM.Value = False
      optM.Value = True
      optInch.Value = False
      optFeet.Value = False
ElseIf gs_PMLunit = "inch" Then
      optMM.Value = False
      optM.Value = False
      optInch.Value = True
      optFeet.Value = False
ElseIf gs_PMLunit = "feet" Then
      optMM.Value = False
      optM.Value = False
      optInch.Value = False
      optFeet.Value = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

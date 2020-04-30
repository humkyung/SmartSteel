VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmSpecHB 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Connection Spec. of Horizontal Bracing"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtCode 
      Height          =   330
      Left            =   1170
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6300
      Width           =   1635
   End
   Begin VB.Frame frmPic 
      Caption         =   "Picture of Type I"
      Height          =   2190
      Left            =   7125
      TabIndex        =   26
      Top             =   3870
      Width           =   1815
      Begin VB.Image imgType 
         Height          =   1725
         Left            =   135
         Top             =   315
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   9390
      TabIndex        =   20
      Top             =   5400
      Width           =   1725
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   465
      Left            =   9390
      TabIndex        =   19
      Top             =   4770
      Width           =   1725
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   9360
      TabIndex        =   18
      Top             =   4140
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Input -> Click ""Save"" Button"
      Height          =   2220
      Left            =   0
      TabIndex        =   2
      Top             =   3870
      Width           =   6810
      Begin MSComctlLib.ImageList ImageList 
         Left            =   720
         Top             =   300
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   105
         ImageHeight     =   115
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHB.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHB.frx":8E46
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdMname 
         Caption         =   "..."
         Height          =   285
         Left            =   6180
         TabIndex        =   34
         Top             =   300
         Width           =   345
      End
      Begin VB.ComboBox txtHname 
         Height          =   300
         Left            =   1620
         TabIndex        =   33
         Text            =   "(Input Data)"
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox txtSpace1 
         Height          =   270
         Left            =   1620
         TabIndex        =   32
         Text            =   "0"
         Top             =   1890
         Width           =   1155
      End
      Begin VB.TextBox txtSpace2 
         Height          =   270
         Left            =   4980
         TabIndex        =   31
         Text            =   "0"
         Top             =   1860
         Width           =   1155
      End
      Begin VB.CommandButton cmdSpace1 
         Caption         =   "..."
         Height          =   285
         Left            =   2790
         TabIndex        =   30
         Top             =   1890
         Width           =   375
      End
      Begin VB.CommandButton cmdSpace2 
         Caption         =   "..."
         Height          =   285
         Left            =   6150
         TabIndex        =   29
         Top             =   1860
         Width           =   375
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   300
         Left            =   1620
         TabIndex        =   25
         Text            =   "mm"
         Top             =   315
         Width           =   1545
      End
      Begin VB.ComboBox cmbShape 
         Height          =   300
         Left            =   1620
         TabIndex        =   22
         Text            =   "Angle"
         Top             =   1575
         Width           =   1545
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   1620
         TabIndex        =   17
         Text            =   "Type I"
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtGage 
         Height          =   270
         Left            =   4995
         TabIndex        =   16
         Text            =   "0"
         Top             =   1575
         Width           =   1545
      End
      Begin VB.TextBox txtPlate 
         Height          =   270
         Left            =   4995
         TabIndex        =   15
         Text            =   "0"
         Top             =   1260
         Width           =   1545
      End
      Begin VB.TextBox txtHspace 
         Height          =   270
         Left            =   4995
         TabIndex        =   14
         Text            =   "0"
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox txtHspaceEA 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4995
         TabIndex        =   13
         Text            =   "0"
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox txtHea 
         Height          =   270
         Left            =   1620
         TabIndex        =   7
         Text            =   "0"
         Top             =   1305
         Width           =   1545
      End
      Begin VB.TextBox txtMName 
         Height          =   270
         Left            =   4995
         TabIndex        =   6
         Text            =   "(Input Data)"
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Space 2               :"
         Height          =   180
         Left            =   3270
         TabIndex        =   28
         Top             =   1920
         Width           =   1650
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Space 1          :"
         Height          =   180
         Left            =   180
         TabIndex        =   27
         Top             =   1950
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Unit                 :"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   405
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Shape Type     :"
         Height          =   180
         Left            =   180
         TabIndex        =   23
         Top             =   1665
         Width           =   1395
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Gage                   :"
         Height          =   180
         Left            =   3285
         TabIndex        =   12
         Top             =   1620
         Width           =   1650
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plate Thk             :"
         Height          =   180
         Left            =   3285
         TabIndex        =   11
         Top             =   1350
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Space of HTB       :"
         Height          =   180
         Left            =   3285
         TabIndex        =   10
         Top             =   1035
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Space EA of HTB  :"
         Height          =   180
         Left            =   3285
         TabIndex        =   9
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EA of HTB        :"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1365
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name of HTB   :"
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   1035
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type               :"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Member Name    :"
         Height          =   180
         Left            =   3330
         TabIndex        =   3
         Top             =   405
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   11220
      Begin FPSpread.vaSpread ssData 
         Height          =   3555
         Left            =   60
         TabIndex        =   35
         Top             =   180
         Width           =   11055
         _Version        =   393216
         _ExtentX        =   19500
         _ExtentY        =   6271
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         OperationMode   =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmSpecHB.frx":11E58
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2190
      Left            =   9225
      TabIndex        =   21
      Top             =   3870
      Width           =   1995
   End
End
Attribute VB_Name = "frmSpecHB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstr_Job As String

Private Sub cmbType_Click()

If CStr(Trim(cmbType.Text)) = "Type I" Then
    Set imgType.Picture = ImageList.ListImages(1).Picture
    frmPic.Caption = "Picture of Type I"
    
Else
    Set imgType.Picture = ImageList.ListImages(2).Picture
    frmPic.Caption = "Picture of Type II"
    
End If

End Sub

Private Sub cmbUnit_Click()
Dim xSQL As String
txtHname.Clear
xSQL = "Select Name from BoltNut where unit = '" & CStr(Trim(cmbUnit.Text)) & "'" & "order by name"

Call Query_AddList2_function(0, txtHname, xSQL)

End Sub

Private Sub cmdDelete_Click()
Dim xSQL As String
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, _
    Temp5 As String, Temp6 As String, Temp7 As String, Temp8 As String, _
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String


Dim Row As Long

Temp1 = gfunSS_GetText(ssData, ssData.ActiveRow, 1)
Temp2 = gfunSS_GetText(ssData, ssData.ActiveRow, 2)
Temp3 = gfunSS_GetText(ssData, ssData.ActiveRow, 3)
Temp4 = gfunSS_GetText(ssData, ssData.ActiveRow, 4)
Temp5 = gfunSS_GetText(ssData, ssData.ActiveRow, 5)
Temp6 = gfunSS_GetText(ssData, ssData.ActiveRow, 6)
Temp7 = gfunSS_GetText(ssData, ssData.ActiveRow, 7)
Temp8 = gfunSS_GetText(ssData, ssData.ActiveRow, 8)
Temp9 = gfunSS_GetText(ssData, ssData.ActiveRow, 9)
Temp10 = gfunSS_GetText(ssData, ssData.ActiveRow, 10)
Temp11 = gfunSS_GetText(ssData, ssData.ActiveRow, 11)
Temp12 = gfunSS_GetText(ssData, ssData.ActiveRow, 12)
Temp13 = gfunSS_GetText(ssData, ssData.ActiveRow, 13)

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               xSQL = "delete from HB_Connection "
               xSQL = xSQL & "where shape = '" & Temp1 & "' "
               xSQL = xSQL & "and member_Name = '" & Temp2 & "' "
               xSQL = xSQL & "and type = '" & Temp3 & "' "
               xSQL = xSQL & "and htb_name = '" & Temp4 & "' "
               xSQL = xSQL & "and htb_num = " & CInt(Trim(Temp5)) & " "
               xSQL = xSQL & "and htb_snum = " & CInt(Trim(Temp6)) & " "
               xSQL = xSQL & "and htb_space = " & CSng(Trim(Temp7)) & " "
               xSQL = xSQL & "and space1 = " & CSng(Trim(Temp8)) & " "
               xSQL = xSQL & "and space2 = " & CSng(Trim(Temp9)) & " "
               xSQL = xSQL & "and plate_thk = " & CSng(Trim(Temp10)) & " "
               xSQL = xSQL & "and gage = " & CSng(Trim(Temp11))
               xSQL = xSQL & "and Unit = '" & CStr(Trim(Temp12)) & "' "
               xSQL = xSQL & "and job = '" & gstr_Job & "' "
               xSQL = xSQL & "and code = '" & CStr(Trim(Temp13)) & "' "
               adoConnection1.Execute (xSQL)
                     
                     
               If ssData.ActiveRow = 1 Then Row = 1
               If ssData.ActiveRow = ssData.MaxRows Then Row = ssData.MaxRows - 1
               
               Call gsubSS_DelRow(ssData, ssData.ActiveRow)
               
               xSQL = "select shape,member_name, type, htb_name, htb_num, htb_snum, htb_space, space1, space2, plate_thk, gage, unit, Code from HB_Connection where job = '" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If

End Sub

Private Sub cmdExit_Click()
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub cmdMname_Click()
    gstr_Shape = CStr(Trim(cmbShape.Text))
    gin_Shape_Flag = 2

    frmMember.Show

End Sub

Private Sub cmdSave_Click()
Dim xSQL As String
Dim Temp As String

If Trim(CStr(cmbType.Text)) = "Type I" Then
    Temp = "I"
Else
    Temp = "II"
End If

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               'shape,member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Gage,Unit
               xSQL = "select Member_Name from HB_Connection where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
               xSQL = xSQL & "and job = '" & gstr_Job & "'"
               If DataCheck(CStr(Trim(txtMname.Text)), xSQL) Then
                   xSQL = "update HB_Connection set "
                   xSQL = xSQL & "Shape ='" & CStr(Trim(cmbShape.Text)) & "', "
                   xSQL = xSQL & "Member_Name ='" & CStr(Trim(txtMname.Text)) & "', "
                   xSQL = xSQL & "type = '" & Temp & "',"
                   xSQL = xSQL & "HTB_Name = '" & CStr(Trim(txtHname.Text)) & "',"
                   xSQL = xSQL & "HTB_Num = " & CSng(Trim(txtHea.Text)) & ","
                   xSQL = xSQL & "HTB_SNum = " & CSng(Trim(txtHspaceEA.Text)) & ","
                   xSQL = xSQL & "HTB_Space = " & CSng(Trim(txtHspace.Text)) & ", "
                   xSQL = xSQL & "Space1 = " & CSng(Trim(txtSpace1.Text)) & ","
                   xSQL = xSQL & "Space2 = " & CSng(Trim(txtSpace2.Text)) & ", "
                   xSQL = xSQL & "Plate_Thk = " & CSng(Trim(txtPlate.Text)) & ", "
                   xSQL = xSQL & "Gage = " & CSng(Trim(txtGage.Text)) & ", "
                   xSQL = xSQL & "Unit ='" & CStr(Trim(cmbUnit.Text)) & "' "
                   xSQL = xSQL & "Where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
                   xSQL = xSQL & "and job = '" & gstr_Job & "'"
                   
                   adoConnection1.Execute (xSQL)
                   
               Else
                   xSQL = "insert into HB_Connection values ('"
                   xSQL = xSQL & gstr_Job & "', '"
                   xSQL = xSQL & Trim(CStr(cmbShape.Text)) & "', '"
                   xSQL = xSQL & Trim(CStr(txtMname.Text)) & "', '"
                   xSQL = xSQL & Temp & "', '"
                   xSQL = xSQL & CStr(Trim(txtHname.Text)) & "',"
                   xSQL = xSQL & CInt(Trim(txtHea.Text)) & ", "
                   xSQL = xSQL & CInt(Trim(txtHspaceEA.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtHspace.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtSpace1.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtSpace2.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtPlate.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtGage.Text)) & ",  '"
                   xSQL = xSQL & CStr(Trim(cmbUnit.Text)) & "', '"
                   xSQL = xSQL & CStr(Trim(txtCode.Text)) & "'"
                   xSQL = xSQL & ")"
                   
                   adoConnection1.Execute (xSQL)
               End If
               
               'Call gsubSS_SetMax(ssData, 0, ssData.MaxCols)
               xSQL = "select shape,member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Space1,Space2,Plate_Thk,Gage,Unit , Code from HB_Connection "
               xSQL = xSQL & "where job = '" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If
End Sub

Private Sub cmdSpace1_Click()
gs_SP_Flag = "3"
frmSpacePicture.Show
End Sub

Private Sub cmdSpace2_Click()
gs_SP_Flag = "4"
frmSpacePicture.Show
End Sub

Private Sub Form_Load()
Dim xSQL As String

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
            Me.Caption = "Current Project : " & gstr_Job & " , Spec. of Horizontal Bracing"
End If

Call gs_project_Input(gstr_Job)

If gstr_Job = "" Then
      MsgBox "Job is not selected. Please Job Select at General"
      End
Else

      xSQL = "select shape,member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Space1,Space2,Plate_Thk,Gage,Unit , code from HB_Connection "
      xSQL = xSQL & "where job = '" & gstr_Job & "' order by member_name"
      
      Call gsubSSADOQuery(1, xSQL, ssData)
 End If
              
      cmbType.AddItem "Type I"
      cmbType.AddItem "Type II"
      
      cmbShape.AddItem "Angle"
      cmbShape.AddItem "Channel"
      cmbShape.AddItem "Double Angle"
      cmbShape.AddItem "Double Channel"
      cmbShape.AddItem "Tee"
      
      cmbUnit.AddItem "mm"
      cmbUnit.AddItem "inch"
      
      'xSQL = "Select Name from BoltNut where unit = '" & gstr_Unit & "'" & "order by name"
      xSQL = "Select Name from BoltNut where unit = 'mm' order by name"
      Call Query_AddList2_function(0, txtHname, xSQL)
      
      Set imgType.Picture = ImageList.ListImages(1).Picture
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub ssData_Click(ByVal Col As Long, ByVal Row As Long)

Dim xSQL As String
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, _
    Temp5 As String, Temp6 As String, Temp7 As String, Temp8 As String, _
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, Temp13 As String

Temp1 = gfunSS_GetText(ssData, ssData.ActiveRow, 1)
Temp2 = gfunSS_GetText(ssData, ssData.ActiveRow, 2)
Temp3 = gfunSS_GetText(ssData, ssData.ActiveRow, 3)
Temp4 = gfunSS_GetText(ssData, ssData.ActiveRow, 4)
Temp5 = gfunSS_GetText(ssData, ssData.ActiveRow, 5)
Temp6 = gfunSS_GetText(ssData, ssData.ActiveRow, 6)
Temp7 = gfunSS_GetText(ssData, ssData.ActiveRow, 7)
Temp8 = gfunSS_GetText(ssData, ssData.ActiveRow, 8)
Temp9 = gfunSS_GetText(ssData, ssData.ActiveRow, 9)
Temp10 = gfunSS_GetText(ssData, ssData.ActiveRow, 10)
Temp11 = gfunSS_GetText(ssData, ssData.ActiveRow, 11)
Temp12 = gfunSS_GetText(ssData, ssData.ActiveRow, 12)
Temp13 = gfunSS_GetText(ssData, ssData.ActiveRow, 13)

cmbShape.Text = Temp1
txtMname.Text = Temp2
If CStr(Trim(Temp3)) = "I" Then
    cmbType.Text = "Type I"
    Set imgType.Picture = ImageList.ListImages(1).Picture
    frmPic.Caption = "Picture of Type I"
Else
    cmbType.Text = "Type II"
    Set imgType.Picture = ImageList.ListImages(2).Picture
    frmPic.Caption = "Picture of Type II"
End If

txtHname.Text = Temp4
txtHea.Text = Temp5
txtHspaceEA.Text = Temp6
txtHspace.Text = Temp7
txtSpace1.Text = Temp8
txtSpace2.Text = Temp9
txtPlate.Text = Temp10
txtGage.Text = Temp11
cmbUnit.Text = Temp12
txtCode.Text = Temp13

End Sub


Private Sub txtHea_LostFocus()
Dim Temp As Integer
Temp = CInt(Trim(txtHea.Text))
If CStr(Trim(cmbType.Text)) = "Type I" Then
               txtHspaceEA.Text = Temp - 1
Else
               txtHspaceEA.Text = Temp / 2 - 1
End If
End Sub


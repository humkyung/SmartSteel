VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmSpecHBP 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Connection Spec. of Hinged Base Plate"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtCode 
      Height          =   270
      Left            =   300
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6720
      Width           =   1860
   End
   Begin VB.Frame frmPic 
      Caption         =   "Picture of Type I"
      Height          =   1995
      Left            =   7020
      TabIndex        =   26
      Top             =   3915
      Width           =   1995
      Begin VB.Image imgBPH 
         Height          =   1545
         Left            =   135
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11445
      Begin MSComctlLib.ImageList ImageList 
         Left            =   60
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   105
         ImageHeight     =   100
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHBP.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHBP.frx":7BC2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHBP.frx":101F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecHBP.frx":18826
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread ssData 
         Height          =   3555
         Left            =   60
         TabIndex        =   33
         Top             =   180
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
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
         SpreadDesigner  =   "frmSpecHBP.frx":20E58
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Input -> Click ""Save"" Button"
      Height          =   1995
      Left            =   0
      TabIndex        =   13
      Top             =   3900
      Width           =   6810
      Begin VB.CommandButton cmdMname 
         Caption         =   "..."
         Height          =   255
         Left            =   6210
         TabIndex        =   32
         Top             =   270
         Width           =   345
      End
      Begin VB.ComboBox txtBname 
         Height          =   300
         Left            =   1620
         TabIndex        =   31
         Text            =   "(Input Data)"
         Top             =   780
         Width           =   1545
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   300
         Left            =   1620
         TabIndex        =   0
         Text            =   "mm"
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtDimB 
         Height          =   270
         Left            =   4995
         TabIndex        =   8
         Text            =   "0"
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtDimD 
         Height          =   270
         Left            =   4995
         TabIndex        =   10
         Text            =   "0"
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txtDimC 
         Height          =   270
         Left            =   1620
         TabIndex        =   9
         Text            =   "0"
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txtDimA 
         Height          =   270
         Left            =   1620
         TabIndex        =   7
         Text            =   "0"
         Top             =   1350
         Width           =   1545
      End
      Begin VB.TextBox txtMName 
         Height          =   270
         Left            =   4995
         TabIndex        =   2
         Text            =   "(Input Data)"
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox txtBea 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1620
         TabIndex        =   6
         Text            =   "0"
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtBPLx 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4995
         TabIndex        =   3
         Text            =   "0"
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtBPLy 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4995
         TabIndex        =   4
         Text            =   "0"
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtBPt 
         Height          =   270
         Left            =   4995
         TabIndex        =   5
         Text            =   "0"
         Top             =   1080
         Width           =   1545
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   1620
         TabIndex        =   1
         Text            =   "Type I"
         Top             =   495
         Width           =   1545
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Select Unit       :"
         Height          =   180
         Left            =   180
         TabIndex        =   30
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Plate Thk             :"
         Height          =   180
         Left            =   3285
         TabIndex        =   29
         Top             =   1125
         Width           =   1635
      End
      Begin VB.Label lblDimB 
         AutoSize        =   -1  'True
         Caption         =   "Dimension B        :"
         Height          =   180
         Left            =   3285
         TabIndex        =   28
         Top             =   1395
         Width           =   1620
      End
      Begin VB.Label lblDimD 
         AutoSize        =   -1  'True
         Caption         =   "Dimension D        :"
         Height          =   180
         Left            =   3285
         TabIndex        =   27
         Top             =   1665
         Width           =   1620
      End
      Begin VB.Label lblDimC 
         AutoSize        =   -1  'True
         Caption         =   "Dimension C    :"
         Height          =   180
         Left            =   180
         TabIndex        =   25
         Top             =   1665
         Width           =   1395
      End
      Begin VB.Label lblDimA 
         AutoSize        =   -1  'True
         Caption         =   "Dimension A    :"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   1395
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Member Name :"
         Height          =   180
         Left            =   3285
         TabIndex        =   20
         Top             =   315
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type               :"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   585
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name of Bolt    :"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EA of Bolt        :"
         Height          =   180
         Left            =   180
         TabIndex        =   17
         Top             =   1125
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X Length of Plate  :"
         Height          =   180
         Left            =   3285
         TabIndex        =   16
         Top             =   585
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y Length of Plate  :"
         Height          =   180
         Left            =   3285
         TabIndex        =   15
         Top             =   855
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   9405
      TabIndex        =   11
      Top             =   4185
      Width           =   1860
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   465
      Left            =   9420
      TabIndex        =   12
      Top             =   4725
      Width           =   1860
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   9405
      TabIndex        =   14
      Top             =   5265
      Width           =   1860
   End
   Begin VB.Frame Frame3 
      Height          =   1995
      Left            =   9225
      TabIndex        =   23
      Top             =   3915
      Width           =   2220
   End
End
Attribute VB_Name = "frmSpecHBP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstr_Job As String

Private Sub cmbType_Click()

If CStr(Trim(cmbType.Text)) = "Type I" Then
    
    Set imgBPH.Picture = ImageList.ListImages(1).Picture

    frmPic.Caption = "Picture of Type I"
    lblDimC.Enabled = True
    lblDimD.Enabled = False
    txtDimC.Enabled = True
    txtDimD.Enabled = False
    txtBea.Text = "2"
ElseIf CStr(Trim(cmbType.Text)) = "Type II" Then
    
    Set imgBPH.Picture = ImageList.ListImages(2).Picture
    frmPic.Caption = "Picture of Type II"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "4"
ElseIf CStr(Trim(cmbType.Text)) = "Type III" Then
    Set imgBPH.Picture = ImageList.ListImages(3).Picture
    frmPic.Caption = "Picture of Type III"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "6"
Else
    Set imgBPH.Picture = ImageList.ListImages(4).Picture
    frmPic.Caption = "Picture of Type IV"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "8"
End If


If CStr(Trim(txtDimA.Text)) = "" Then
      txtDimA.Text = "0"
End If
If CStr(Trim(txtDimB.Text)) = "" Then
      txtDimA.Text = "0"
End If
txtBPLx.Text = CStr(CSng(Trim(txtDimA.Text)) + 2 * CSng(Trim(txtDimB.Text)))
If CStr(Trim(txtDimC.Text)) = "" Then
      txtDimC.Text = "0"
End If
If CStr(Trim(txtDimD.Text)) = "" Then
      txtDimD.Text = "0"
End If
If CStr(Trim(cmbType.Text)) = "Type I" Then
      txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)))
ElseIf CStr(Trim(cmbType.Text)) = "Type II" Then
      txtBPLy.Text = CStr(CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
ElseIf CStr(Trim(cmbType.Text)) = "Type III" Then
      txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
Else
      txtBPLy.Text = CStr(3 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
End If
End Sub

Private Sub cmbUnit_Click()
Dim xSQL As String
txtBname.Clear
xSQL = "Select Name from BoltNut where unit = '" & CStr(Trim(cmbUnit.Text)) & "'" & "order by name"

Call Query_AddList2_function(0, txtBname, xSQL)

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

               xSQL = "delete from BasePlate_Hinged where job = '" & gstr_Job & "' "
               xSQL = xSQL & "and Member_Name = '" & CStr(Trim(Temp1)) & "' "
               xSQL = xSQL & "and Xlen = " & CSng(Trim(Temp2)) & " "
               xSQL = xSQL & "and Ylen = " & CSng(Trim(Temp3)) & " "
               xSQL = xSQL & "and Pthk = " & CSng(Trim(Temp4)) & " "
               xSQL = xSQL & "and Xbtob = " & CSng(Trim(Temp5)) & " "
               xSQL = xSQL & "and Xcls = " & CSng(Trim(Temp6)) & " "
               xSQL = xSQL & "and Ybtob = " & CSng(Trim(Temp7)) & " "
               xSQL = xSQL & "and Ycls = " & CSng(Trim(Temp8)) & " "
               xSQL = xSQL & "and BoltEA = " & CInt(Trim(Temp9)) & " "
               xSQL = xSQL & "and BoltName = '" & CStr(Trim(Temp10)) & "' "
               xSQL = xSQL & "and Type = '" & CStr(Trim(Temp11)) & "' "
               xSQL = xSQL & "and Unit = '" & CStr(Trim(Temp12)) & "' "
               xSQL = xSQL & "and code = '" & CStr(Trim(Temp13)) & "' "
               
               adoConnection1.Execute (xSQL)
                     
                     
               If ssData.ActiveRow = 1 Then Row = 1
               If ssData.ActiveRow = ssData.MaxRows Then Row = ssData.MaxRows - 1
               
               Call gsubSS_DelRow(ssData, ssData.ActiveRow)
               
               xSQL = "select member_name, xlen, ylen, pthk, xbtob, xcls, ybtob, ycls, boltea, boltname, type, unit , code from BasePlate_Hinged where job = '" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If

End Sub

Private Sub cmdExit_Click()
'For i = CInt(Trim(Me.Height)) To 0 Step -60
'               Me.Width = i
'               Me.Height = i
'Next i

If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub cmdMname_Click()
    gstr_Shape = "Hbeam"
    gin_Shape_Flag = 3
    frmMember.Show

End Sub

Private Sub cmdSave_Click()
Dim xSQL As String
Dim Temp As String

If Trim(CStr(cmbType.Text)) = "Type I" Then
    Temp = "I"
ElseIf Trim(CStr(cmbType.Text)) = "Type II" Then
    Temp = "II"
ElseIf Trim(CStr(cmbType.Text)) = "Type III" Then
    Temp = "III"
Else
    Temp = "IV"
End If

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               xSQL = "select Member_Name from BasePlate_Hinged where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
               xSQL = xSQL & "and job = '" & gstr_Job & "'"
               If DataCheck(CStr(Trim(txtMname.Text)), xSQL) Then
                   xSQL = "update BasePlate_Hinged set "
                   xSQL = xSQL & "Member_Name ='" & CStr(Trim(txtMname.Text)) & "', "
                   xSQL = xSQL & "Xlen = " & CSng(Trim(txtBPLx.Text)) & ","
                   xSQL = xSQL & "Ylen = " & CSng(Trim(txtBPLy.Text)) & ","
                   xSQL = xSQL & "Pthk = " & CSng(Trim(txtBPt.Text)) & ","
                   xSQL = xSQL & "Xbtob = " & CSng(Trim(txtDimA.Text)) & ", "
                   xSQL = xSQL & "Xcls = " & CSng(Trim(txtDimB.Text)) & ", "
                   xSQL = xSQL & "Ybtob = " & CSng(Trim(txtDimC.Text)) & ", "
                   xSQL = xSQL & "Ycls = " & CSng(Trim(txtDimD.Text)) & ", "
                   xSQL = xSQL & "BoltEA = " & CInt(Trim(txtBea.Text)) & ", "
                   xSQL = xSQL & "BoltName ='" & CStr(Trim(txtBname.Text)) & "', "
                   xSQL = xSQL & "Type ='" & Temp & "' "
                   xSQL = xSQL & "Where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
                   xSQL = xSQL & "and job = '" & gstr_Job & "'"
                   adoConnection1.Execute (xSQL)
                   
               Else
                   xSQL = "insert into BasePlate_Hinged values ('"
                   xSQL = xSQL & gstr_Job & "', '"
                   xSQL = xSQL & CStr(Trim(txtMname.Text)) & "',"
                   xSQL = xSQL & CSng(Trim(txtBPLx.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtBPLy.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtBPt.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtDimA.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtDimB.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtDimC.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtDimD.Text)) & ","
                   xSQL = xSQL & CInt(Trim(txtBea.Text)) & ", '"
                   xSQL = xSQL & CStr(Trim(txtBname.Text)) & "', '"
                   xSQL = xSQL & Temp & "', '"
                   xSQL = xSQL & CStr(Trim(cmbUnit.Text)) & "', '"
                   xSQL = xSQL & CStr(Trim(txtCode.Text)) & "')"
                   adoConnection1.Execute (xSQL)
               
               End If
               
               xSQL = "select Member_Name,Xlen,Ylen,Pthk,Xbtob,Xcls,Ybtob,Ycls,BoltEA,BoltName,Type,unit , code from BasePlate_Hinged "
               xSQL = xSQL & "where job = '" & gstr_Job & "'"
               Call gsubSSADOQuery(1, xSQL, ssData)
End If
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
            Me.Caption = "Current Project : " & gstr_Job & " , Spec. of Hinged Base Plate"
End If


Call gs_project_Input(gstr_Job)

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
Else
               xSQL = "select Member_Name,Xlen,Ylen,Pthk,Xbtob,Xcls,Ybtob,Ycls,BoltEA,BoltName,Type,unit, Code from BasePlate_Hinged "
               xSQL = xSQL & "where job = '" & gstr_Job & "' order by member_name"
               Call gsubSSADOQuery(1, xSQL, ssData)
End If

cmbType.AddItem "Type I"
cmbType.AddItem "Type II"
cmbType.AddItem "Type III"
cmbType.AddItem "Type IV"
txtBea.Text = "2"

cmbUnit.AddItem "mm"
cmbUnit.AddItem "inch"

Set imgBPH.Picture = ImageList.ListImages(1).Picture

lblDimC.Enabled = True
lblDimD.Enabled = False
txtDimD.Enabled = False

xSQL = "Select Name from BoltNut where unit = 'mm' order by name"
Call Query_AddList2_function(0, txtBname, xSQL)

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

txtMname.Text = Temp1
txtBPLx.Text = Temp2
txtBPLy.Text = Temp3
txtBPt.Text = Temp4
txtDimA.Text = Temp5
txtDimB.Text = Temp6
txtDimC.Text = Temp7
txtDimD.Text = Temp8
txtBea.Text = Temp9
txtBname.Text = Temp10
txtCode.Text = Temp13

If CStr(Trim(Temp11)) = "I" Then
    cmbType.Text = "Type I"
    Set imgBPH.Picture = ImageList.ListImages(1).Picture

    frmPic.Caption = "Picture of Type I"
    lblDimC.Enabled = True
    lblDimD.Enabled = False
    txtDimC.Enabled = True
    txtDimD.Enabled = False
    txtBea.Text = "2"
ElseIf CStr(Trim(Temp11)) = "II" Then
    cmbType.Text = "Type II"
    Set imgBPH.Picture = ImageList.ListImages(2).Picture

    frmPic.Caption = "Picture of Type II"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "4"
ElseIf CStr(Trim(Temp11)) = "III" Then
    cmbType.Text = "Type III"
    Set imgBPH.Picture = ImageList.ListImages(3).Picture

    frmPic.Caption = "Picture of Type III"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "6"
Else
    cmbType.Text = "Type IV"
    Set imgBPH.Picture = ImageList.ListImages(4).Picture

    frmPic.Caption = "Picture of Type IV"
    lblDimC.Enabled = True
    lblDimD.Enabled = True
    txtDimC.Enabled = True
    txtDimD.Enabled = True
    txtBea.Text = "8"
End If
cmbUnit.Text = Temp12


If CStr(Trim(txtDimA.Text)) = "" Then
      txtDimA.Text = "0"
End If
If CStr(Trim(txtDimB.Text)) = "" Then
      txtDimA.Text = "0"
End If
txtBPLx.Text = CStr(CSng(Trim(txtDimA.Text)) + 2 * CSng(Trim(txtDimB.Text)))
If CStr(Trim(txtDimC.Text)) = "" Then
      txtDimC.Text = "0"
End If
If CStr(Trim(txtDimD.Text)) = "" Then
      txtDimD.Text = "0"
End If
If CStr(Trim(cmbType.Text)) = "Type I" Then
      txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)))
ElseIf CStr(Trim(cmbType.Text)) = "Type II" Then
      txtBPLy.Text = CStr(CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
ElseIf CStr(Trim(cmbType.Text)) = "Type III" Then
      txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
Else
      txtBPLy.Text = CStr(3 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
End If



End Sub

Private Sub txtDimA_Change()
      If CStr(Trim(txtDimA.Text)) = "" Then
            txtDimA.Text = "0"
      End If
      If CStr(Trim(txtDimB.Text)) = "" Then
            txtDimB.Text = "0"
      End If
      txtBPLx.Text = CStr(CSng(Trim(txtDimA.Text)) + 2 * CSng(Trim(txtDimB.Text)))
End Sub

Private Sub txtDimB_Change()
      If CStr(Trim(txtDimA.Text)) = "" Then
            txtDimA.Text = "0"
      End If
      If CStr(Trim(txtDimB.Text)) = "" Then
            txtDimB.Text = "0"
      End If
      txtBPLx.Text = CStr(CSng(Trim(txtDimA.Text)) + 2 * CSng(Trim(txtDimB.Text)))
End Sub

Private Sub txtDimC_Change()
      If CStr(Trim(txtDimC.Text)) = "" Then
            txtDimC.Text = "0"
      End If
      If CStr(Trim(txtDimD.Text)) = "" Then
            txtDimD.Text = "0"
      End If
      If CStr(Trim(cmbType.Text)) = "Type I" Then
            txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)))
      ElseIf CStr(Trim(cmbType.Text)) = "Type II" Then
            txtBPLy.Text = CStr(CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      ElseIf CStr(Trim(cmbType.Text)) = "Type III" Then
            txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      Else
            txtBPLy.Text = CStr(3 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      End If
End Sub

Private Sub txtDimD_Change()
      If CStr(Trim(txtDimC.Text)) = "" Then
            txtDimC.Text = "0"
      End If
      If CStr(Trim(txtDimD.Text)) = "" Then
            txtDimD.Text = "0"
      End If
      If CStr(Trim(cmbType.Text)) = "Type I" Then
            txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)))
      ElseIf CStr(Trim(cmbType.Text)) = "Type II" Then
            txtBPLy.Text = CStr(CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      ElseIf CStr(Trim(cmbType.Text)) = "Type III" Then
            txtBPLy.Text = CStr(2 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      Else
            txtBPLy.Text = CStr(3 * CSng(Trim(txtDimC.Text)) + 2 * CSng(Trim(txtDimD.Text)))
      End If
End Sub

VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmSpecSC 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Connection Spec of Shear"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtCode 
      Height          =   510
      Left            =   315
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6210
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data Input -> Click ""Save"" Button"
      Height          =   2010
      Left            =   0
      TabIndex        =   2
      Top             =   3870
      Width           =   8790
      Begin VB.CommandButton cmdMname 
         Caption         =   "...."
         Height          =   300
         Left            =   2760
         TabIndex        =   38
         Top             =   960
         Width           =   390
      End
      Begin VB.ComboBox txtHname 
         Height          =   300
         Left            =   1620
         TabIndex        =   37
         Text            =   "(Input Data)"
         Top             =   1290
         Width           =   1545
      End
      Begin VB.TextBox txtHspaceEA 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4920
         TabIndex        =   34
         Text            =   "0"
         Top             =   660
         Width           =   1155
      End
      Begin VB.TextBox txtHspace 
         Height          =   270
         Left            =   4650
         TabIndex        =   33
         Text            =   "0"
         Top             =   975
         Width           =   1425
      End
      Begin VB.TextBox txtStiffTHK 
         Height          =   270
         Left            =   4650
         TabIndex        =   32
         Text            =   "0"
         Top             =   1590
         Width           =   1425
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "...."
         Height          =   300
         Left            =   2745
         TabIndex        =   30
         Top             =   630
         Width           =   390
      End
      Begin VB.TextBox txtEvalue 
         Height          =   270
         Left            =   7155
         TabIndex        =   29
         Text            =   "0"
         Top             =   1545
         Width           =   1560
      End
      Begin VB.TextBox txtHea 
         Height          =   270
         Left            =   1620
         TabIndex        =   27
         Text            =   "0"
         Top             =   1620
         Width           =   1515
      End
      Begin VB.TextBox txtMname 
         Height          =   270
         Left            =   1620
         TabIndex        =   11
         Text            =   "(Input Data)"
         Top             =   990
         Width           =   1125
      End
      Begin VB.TextBox txtGap 
         Height          =   270
         Left            =   4650
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtAvalue 
         Height          =   270
         Left            =   7170
         TabIndex        =   9
         Text            =   "0"
         Top             =   345
         Width           =   1545
      End
      Begin VB.TextBox txtBvalue 
         Height          =   270
         Left            =   7170
         TabIndex        =   8
         Text            =   "0"
         Top             =   660
         Width           =   1545
      End
      Begin VB.TextBox txtCvalue 
         Height          =   270
         Left            =   7170
         TabIndex        =   7
         Text            =   "0"
         Top             =   975
         Width           =   1545
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   1620
         TabIndex        =   6
         Text            =   "Type I"
         Top             =   630
         Width           =   1095
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   300
         Left            =   1620
         TabIndex        =   5
         Text            =   "mm"
         Top             =   315
         Width           =   1545
      End
      Begin VB.TextBox txtDvalue 
         Height          =   270
         Left            =   7155
         TabIndex        =   4
         Text            =   "0"
         Top             =   1260
         Width           =   1560
      End
      Begin VB.TextBox txtPlateTHK 
         Height          =   270
         Left            =   4650
         TabIndex        =   3
         Text            =   "0"
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Space EA of HTB  :"
         Height          =   180
         Left            =   3210
         TabIndex        =   36
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Space of HTB  :"
         Height          =   180
         Left            =   3210
         TabIndex        =   35
         Top             =   1065
         Width           =   1350
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Stiff Plate THK :"
         Height          =   180
         Left            =   3210
         TabIndex        =   31
         Top             =   1635
         Width           =   1320
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "E Value :"
         Height          =   180
         Left            =   6285
         TabIndex        =   28
         Top             =   1590
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type               :"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Member Name :"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   1035
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name of HTB   :"
         Height          =   180
         Left            =   180
         TabIndex        =   20
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gap           :"
         Height          =   180
         Left            =   3210
         TabIndex        =   19
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "A Value : "
         Height          =   180
         Left            =   6300
         TabIndex        =   18
         Top             =   435
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "B Value :"
         Height          =   180
         Left            =   6300
         TabIndex        =   17
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "C Value :"
         Height          =   180
         Left            =   6300
         TabIndex        =   16
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "EA of HTB        :"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   1665
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Unit                 :"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   405
         Width           =   1395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Plate THK        :"
         Height          =   180
         Left            =   3210
         TabIndex        =   13
         Top             =   1350
         Width           =   1380
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "D Value :"
         Height          =   180
         Left            =   6300
         TabIndex        =   12
         Top             =   1320
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   15
      TabIndex        =   0
      Top             =   -45
      Width           =   10575
      Begin FPSpread.vaSpread ssData 
         Height          =   3615
         Left            =   60
         TabIndex        =   39
         Top             =   180
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
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
         MaxCols         =   15
         OperationMode   =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmSpecSC.frx":0000
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   9000
      TabIndex        =   23
      Top             =   5250
      Width           =   1500
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   465
      Left            =   9000
      TabIndex        =   24
      Top             =   4710
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   9000
      TabIndex        =   25
      Top             =   4170
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   1995
      Left            =   8865
      TabIndex        =   26
      Top             =   3870
      Width           =   1770
   End
End
Attribute VB_Name = "frmSpecSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gstr_Job As String
'Private Sub cmbType_Click()
'
'Select Case CStr(Trim(cmbType.Text))
'    Case "Type A1"
'        gs_SP_Flag = "A1"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = False
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = False
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type A2"
'        gs_SP_Flag = "A2"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = False
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type A3"
'        gs_SP_Flag = "A3"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = False
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type A4"
'        gs_SP_Flag = "A4"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = False
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = True
'        txtJvalue.Enabled = True
'    Case "Type B1"
'        gs_SP_Flag = "B1"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = True
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = False
'        txtGvalue.Enabled = False
'        txtHvalue.Enabled = False
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type B2"
'        gs_SP_Flag = "B2"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = True
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = False
'        txtHvalue.Enabled = False
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type B3"
'        gs_SP_Flag = "B3"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = True
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = False
'        txtJvalue.Enabled = False
'    Case "Type B4"
'        gs_SP_Flag = "B4"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = True
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = True
'        txtJvalue.Enabled = False
'    Case "Type B5"
'        gs_SP_Flag = "B5"
'
'        txtLvalue.Enabled = True
'        txtL2value.Enabled = True
'        txtWvalue.Enabled = True
'        txtBvalue.Enabled = True
'        txtCvalue.Enabled = True
'        txtDvalue.Enabled = True
'        txtEvalue.Enabled = True
'        txtFvalue.Enabled = True
'        txtGvalue.Enabled = True
'        txtHvalue.Enabled = True
'        txtIvalue.Enabled = True
'        txtJvalue.Enabled = True
'End Select
'
'End Sub

Private Sub cmbType_Click()
Select Case CStr(Trim(cmbType.Text))
    Case "Type I"
        gs_SP_Flag = "I"
        txtEvalue.Enabled = False
    Case "Type II"
        gs_SP_Flag = "II"
        txtEvalue.Enabled = True
End Select
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
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, _
    Temp13 As String, Temp14 As String, Temp15 As String, Temp16 As String


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
Temp14 = gfunSS_GetText(ssData, ssData.ActiveRow, 14)
Temp15 = gfunSS_GetText(ssData, ssData.ActiveRow, 15)
Temp16 = gfunSS_GetText(ssData, ssData.ActiveRow, 16)
If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               xSQL = "delete from SC_Connection "
               xSQL = xSQL & "where member_Name = '" & Temp1 & "' "
               xSQL = xSQL & "and type = '" & Temp2 & "' "
               xSQL = xSQL & "and htb_name = '" & Temp3 & "' "
               xSQL = xSQL & "and htb_num = " & gf_StringtoSingle(Trim(Temp4)) & " "
               xSQL = xSQL & "and htb_snum = " & gf_StringtoSingle(Trim(Temp5)) & " "
               xSQL = xSQL & "and htb_space = " & gf_StringtoSingle(Trim(Temp6)) & " "
               xSQL = xSQL & "and plate_thk = " & gf_StringtoSingle(Trim(Temp7)) & " "
               xSQL = xSQL & "and stiff_thk = " & gf_StringtoSingle(Trim(Temp8)) & " "
               xSQL = xSQL & "and Gap = " & gf_StringtoSingle(Trim(Temp9)) & " "
               xSQL = xSQL & "and A = " & gf_StringtoSingle(Trim(Temp10)) & " "
               xSQL = xSQL & "and B = " & gf_StringtoSingle(Trim(Temp11)) & " "
               xSQL = xSQL & "and C = " & gf_StringtoSingle(Trim(Temp12)) & " "
               xSQL = xSQL & "and D = " & gf_StringtoSingle(Trim(Temp13)) & " "
               xSQL = xSQL & "and E = " & gf_StringtoSingle(Trim(Temp14)) & " "
               xSQL = xSQL & "and Unit = '" & Trim(Temp15) & "' "
               xSQL = xSQL & "and code = '" & Trim(Temp16) & "' "
               xSQL = xSQL & "and job = '" & gstr_Job & "'"
               
               adoConnection1.Execute (xSQL)
                     
               If ssData.ActiveRow = 1 Then Row = 1
               If ssData.ActiveRow = ssData.MaxRows Then Row = ssData.MaxRows - 1
               
               Call gsubSS_DelRow(ssData, ssData.ActiveRow)
               
               xSQL = "select member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Stiff_Thk," & _
                           "Gap,A,B,C,D,E,Unit,Code from SC_Connection where job = '" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If
End Sub

Private Sub cmdExit_Click()
'For i = CInt(Trim(Me.Height)) To 0 Step -1
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
gin_Shape_Flag = 5
frmMember.Show

End Sub

Private Sub cmdSave_Click()
Dim xSQL As String
Dim Temp As String

Select Case Trim(CStr(cmbType.Text))
    Case "Type I"
        Temp = "I"
    Case "Type II"
        Temp = "II"
End Select
If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               'shape,member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Gage,Unit
               xSQL = "select Member_Name from SC_Connection where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
               xSQL = xSQL & "and job = '" & gstr_Job & "'"
               If DataCheck(CStr(Trim(txtMname.Text)), xSQL) Then
                   xSQL = "update SC_Connection set "
                   xSQL = xSQL & "Member_Name ='" & CStr(Trim(txtMname.Text)) & "', "
                   xSQL = xSQL & "type = '" & Temp & "',"
                   xSQL = xSQL & "HTB_Name = '" & CStr(Trim(txtHname.Text)) & "',"
                   xSQL = xSQL & "HTB_Num = " & CSng(Trim(txtHea.Text)) & ", "
                   xSQL = xSQL & "HTB_SNum = " & CSng(Trim(txtHspaceEA.Text)) & ","
                   xSQL = xSQL & "HTB_Space = " & CSng(Trim(txtHspace.Text)) & ", "
                   xSQL = xSQL & "Plate_Thk = " & CSng(Trim(txtPlateTHK.Text)) & ", "
                   xSQL = xSQL & "stiff_Thk = " & CSng(Trim(txtStiffTHK.Text)) & ", "
                   xSQL = xSQL & "Gap = " & CSng(Trim(txtGap.Text)) & ", "
                   xSQL = xSQL & "A = " & CSng(Trim(txtAvalue.Text)) & ", "
                   xSQL = xSQL & "B = " & CSng(Trim(txtBvalue.Text)) & ", "
                   xSQL = xSQL & "C = " & CSng(Trim(txtCvalue.Text)) & ", "
                   xSQL = xSQL & "D = " & CSng(Trim(txtDvalue.Text)) & ", "
                   xSQL = xSQL & "E = " & CSng(Trim(txtEvalue.Text)) & ", "
                   xSQL = xSQL & "Unit ='" & CStr(Trim(cmbUnit.Text)) & "' "
                   xSQL = xSQL & "Where Member_Name = '" & CStr(Trim(txtMname.Text)) & "' "
                   xSQL = xSQL & "and job = '" & gstr_Job & "'"
                   adoConnection1.Execute (xSQL)
                   
               Else
                   xSQL = "insert into SC_Connection values ("
                   xSQL = xSQL & "'" & gstr_Job & "', "
                   xSQL = xSQL & "'" & Trim(CStr(txtMname.Text)) & "', "
                   xSQL = xSQL & "'" & Temp & "', "
                   xSQL = xSQL & "'" & CStr(Trim(txtHname.Text)) & "', "
                   xSQL = xSQL & CInt(Trim(txtHea.Text)) & ", "
                   xSQL = xSQL & CInt(Trim(txtHspaceEA.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtHspace.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtPlateTHK.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtStiffTHK.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtGap.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtAvalue.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtBvalue.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtCvalue.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtDvalue.Text)) & ", "
                   xSQL = xSQL & CSng(Trim(txtEvalue.Text)) & ", "
                   xSQL = xSQL & "'" & CStr(Trim(cmbUnit.Text)) & "', "
                   xSQL = xSQL & "'" & CStr(Trim(txtCode.Text)) & "' "
                   xSQL = xSQL & ")"
                   
                   adoConnection1.Execute (xSQL)
               End If
               
               'Call gsubSS_SetMax(ssData, 0, ssData.MaxCols)
               xSQL = "select member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Stiff_Thk," & _
                           "Gap,A,B,C,D,E,Unit ,Code from SC_Connection where job = '" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If

End Sub

Private Sub cmdSpace1_Click()

gs_SP_Flag = "1"
frmSpacePicture.Show

End Sub

Private Sub cmdSpace2_Click()
gs_SP_Flag = "2"
frmSpacePicture.Show
End Sub

Private Sub cmdView_Click()
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
            Me.Caption = "Current Project : " & gstr_Job & " , Spec. of Shear Connection"
End If

Call gs_project_Input(gstr_Job)

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
Else

               xSQL = "select member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Stiff_Thk," & _
                             "Gap,A,B,C,D,E,Unit, Code from SC_Connection where job = '" & gstr_Job & "' order by member_name"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If
cmbType.AddItem "Type I"
cmbType.AddItem "Type II"

Call cmbType_Click


'cmbShape.AddItem "Angle"
'cmbShape.AddItem "Channel"
'cmbShape.AddItem "Double Angle"
'cmbShape.AddItem "Double Channel"
'cmbShape.AddItem "Tee"

cmbUnit.AddItem "mm"
cmbUnit.AddItem "inch"
txtEvalue.Enabled = False

xSQL = "Select Name from BoltNut where unit = 'mm' order by name"
Call Query_AddList2_function(0, txtHname, xSQL)

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
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, _
    Temp13 As String, Temp14 As String, Temp15 As String, Temp16 As String

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
Temp14 = gfunSS_GetText(ssData, ssData.ActiveRow, 14)
Temp15 = gfunSS_GetText(ssData, ssData.ActiveRow, 15)
Temp16 = gfunSS_GetText(ssData, ssData.ActiveRow, 16)
txtMname.Text = Temp1
Select Case Temp2
    Case "I"
        cmbType.Text = "Type I"
        txtEvalue.Enabled = False
        gs_SP_Flag = "I"
    Case "II"
        cmbType.Text = "Type II"
        txtEvalue.Enabled = True
        gs_SP_Flag = "II"
End Select
txtHname.Text = Temp3
txtHea.Text = Temp4
txtHspaceEA.Text = Temp5
txtHspace.Text = Temp6
txtPlateTHK.Text = Temp7
txtStiffTHK.Text = Temp8
txtGap.Text = Temp9
txtAvalue.Text = Temp10
txtBvalue.Text = Temp11
txtCvalue.Text = Temp12
txtDvalue.Text = Temp13
txtEvalue.Text = Temp14
cmbUnit.Text = Temp15
txtCode.Text = Temp16

End Sub


'Private Sub txtHea_LostFocus()
'Dim Temp As Integer
'
'Temp = CInt(Trim(txtHea.Text))
'txtHspaceEA.Text = Temp - 1
'End Sub


Private Sub txtHea_LostFocus()
If CStr(Trim(cmbType.Text)) = "Type I" Then
               txtHspaceEA.Text = CInt(Trim(txtHea.Text)) - 1
Else
               txtHspaceEA.Text = CInt(Trim(txtHea.Text)) / 2 - 1
End If

End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmBasePlate_Fixed 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   Icon            =   "frmBasePlate_Fixed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10950
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5220
      Top             =   2820
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
            Picture         =   "frmBasePlate_Fixed.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":1A3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4500
      Top             =   2820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   151
      ImageHeight     =   179
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":3444
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":1736E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":2B298
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":3F1C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":530EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":67016
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":7AF40
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBasePlate_Fixed.frx":8EE6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkNut 
      Caption         =   "Nut Model Check (One Nut)"
      Height          =   360
      Left            =   5340
      TabIndex        =   54
      Top             =   5610
      Value           =   1  '확인
      Width           =   2190
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9435
      TabIndex        =   31
      Top             =   5580
      Width           =   1500
   End
   Begin VB.CommandButton cmbMake 
      Caption         =   "Make PML"
      Height          =   495
      Left            =   7860
      TabIndex        =   34
      Top             =   5580
      Width           =   1530
   End
   Begin VB.Frame Frame7 
      Caption         =   "Select Column Member"
      Height          =   2130
      Left            =   2475
      TabIndex        =   33
      Top             =   3915
      Width           =   2805
      Begin VB.ComboBox cmbColumn 
         Height          =   300
         Left            =   90
         TabIndex        =   57
         Top             =   630
         Width           =   2655
      End
      Begin VB.ListBox lstColumn 
         Height          =   960
         Left            =   90
         TabIndex        =   40
         Top             =   960
         Width           =   2625
      End
      Begin VB.ComboBox cmbCode 
         Height          =   300
         Left            =   765
         TabIndex        =   35
         Text            =   "JIS"
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Code "
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Rib Plate Feature"
      Height          =   2040
      Left            =   0
      TabIndex        =   32
      Top             =   3015
      Width           =   2445
      Begin VB.Image Image3 
         Height          =   1560
         Left            =   90
         Top             =   315
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   450
      Left            =   7920
      TabIndex        =   30
      Top             =   4965
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save D/B"
      Height          =   450
      Left            =   7905
      TabIndex        =   29
      Top             =   4425
      Width           =   2895
   End
   Begin VB.Frame Frame5 
      Caption         =   "Select PML Unit"
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   5085
      Width           =   2445
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   1305
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   360
         TabIndex        =   8
         Top             =   540
         Width           =   690
      End
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   1305
         TabIndex        =   7
         Top             =   540
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Base Plate Data"
      Height          =   5460
      Left            =   7785
      TabIndex        =   5
      Top             =   60
      Width           =   3120
      Begin VB.CheckBox chkI 
         Caption         =   "Check1"
         Height          =   195
         Left            =   780
         TabIndex        =   59
         Top             =   3000
         Width           =   195
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   135
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   4500
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox txtRBT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   55
         Text            =   "0"
         Top             =   3750
         Width           =   1545
      End
      Begin VB.TextBox txtBname 
         Height          =   270
         Left            =   1725
         TabIndex        =   52
         Text            =   "(Input Data)"
         Top             =   4050
         Width           =   1305
      End
      Begin VB.TextBox txtBPT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   50
         Text            =   "0"
         Top             =   3465
         Width           =   1545
      End
      Begin VB.TextBox txtRDN 
         Height          =   270
         Left            =   90
         TabIndex        =   49
         Text            =   "AC-15"
         Top             =   5070
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txtDM 
         Height          =   270
         Left            =   90
         TabIndex        =   47
         Text            =   "A"
         Top             =   4530
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox txtIfactor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   990
         TabIndex        =   44
         Text            =   "1.5"
         Top             =   2925
         Width           =   600
      End
      Begin VB.TextBox txtJ 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   39
         Text            =   "0"
         Top             =   3195
         Width           =   1545
      End
      Begin VB.TextBox txtI 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   270
         Left            =   2250
         TabIndex        =   38
         Text            =   "0"
         Top             =   2925
         Width           =   780
      End
      Begin VB.TextBox txtH 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   270
         Left            =   1485
         TabIndex        =   37
         Text            =   "0"
         Top             =   2655
         Width           =   1545
      End
      Begin VB.ComboBox cmbUnit 
         Height          =   300
         Left            =   1485
         TabIndex        =   28
         Text            =   "mm"
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1485
         TabIndex        =   27
         Top             =   495
         Width           =   1545
      End
      Begin VB.TextBox txtF 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   26
         Text            =   "0"
         Top             =   2115
         Width           =   1545
      End
      Begin VB.TextBox txtD 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   25
         Text            =   "0"
         Top             =   1575
         Width           =   1545
      End
      Begin VB.TextBox txtG 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   24
         Text            =   "0"
         Top             =   2385
         Width           =   1545
      End
      Begin VB.TextBox txtE 
         Height          =   270
         Left            =   1485
         TabIndex        =   23
         Text            =   "0"
         Top             =   1845
         Width           =   1545
      End
      Begin VB.TextBox txtC 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   22
         Text            =   "0"
         Top             =   1305
         Width           =   1545
      End
      Begin VB.TextBox txtB 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   21
         Text            =   "0"
         Top             =   1035
         Width           =   1545
      End
      Begin VB.TextBox txtA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   1
         EndProperty
         Height          =   270
         Left            =   1485
         TabIndex        =   20
         Text            =   "0"
         Top             =   765
         Width           =   1545
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Rib Plate Thk"
         Height          =   180
         Left            =   90
         TabIndex        =   56
         Top             =   3795
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Name of Bolt    :"
         Height          =   180
         Left            =   90
         TabIndex        =   53
         Top             =   4065
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Base Plate Thk"
         Height          =   180
         Left            =   90
         TabIndex        =   51
         Top             =   3510
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Reference DWG No :"
         Height          =   180
         Left            =   45
         TabIndex        =   48
         Top             =   4845
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Detail Mark :"
         Height          =   180
         Left            =   90
         TabIndex        =   46
         Top             =   4305
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "x H ="
         Height          =   180
         Left            =   1665
         TabIndex        =   45
         Top             =   2970
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "J Value"
         Height          =   180
         Left            =   135
         TabIndex        =   43
         Top             =   3240
         Width           =   630
      End
      Begin VB.Label Label14 
         Caption         =   "I Value"
         Height          =   150
         Left            =   135
         TabIndex        =   42
         Top             =   2970
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "H Value"
         Height          =   180
         Left            =   135
         TabIndex        =   41
         Top             =   2700
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "G value"
         Height          =   180
         Left            =   135
         TabIndex        =   19
         Top             =   2430
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "F value"
         Height          =   180
         Left            =   135
         TabIndex        =   18
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "E value"
         Height          =   180
         Left            =   135
         TabIndex        =   17
         Top             =   1890
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "D value"
         Height          =   180
         Left            =   135
         TabIndex        =   16
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C value"
         Height          =   180
         Left            =   135
         TabIndex        =   15
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "B value"
         Height          =   180
         Left            =   135
         TabIndex        =   14
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "A value"
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Column Name"
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unit"
         Height          =   180
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "View Saved Data"
      Height          =   3840
      Left            =   2475
      TabIndex        =   4
      Top             =   45
      Width           =   5280
      Begin FPSpread.vaSpread ssData 
         Height          =   3615
         Left            =   120
         TabIndex        =   60
         Top             =   180
         Width           =   5055
         _Version        =   393216
         _ExtentX        =   8916
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         SpreadDesigner  =   "frmBasePlate_Fixed.frx":9A4BC
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shape Option"
      Height          =   1605
      Left            =   5310
      TabIndex        =   1
      Top             =   3915
      Width           =   2445
      Begin VB.OptionButton Opt01 
         Caption         =   "Option1"
         Height          =   540
         Left            =   135
         TabIndex        =   3
         Top             =   495
         Value           =   -1  'True
         Width           =   285
      End
      Begin VB.OptionButton Opt02 
         Caption         =   "Option2"
         Height          =   540
         Left            =   1200
         TabIndex        =   2
         Top             =   495
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   450
         Top             =   495
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   1575
         Top             =   495
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base Plate Feature"
      Height          =   2940
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   2445
      Begin VB.Image imgFeature 
         Height          =   2685
         Left            =   90
         Picture         =   "frmBasePlate_Fixed.frx":9C299
         Top             =   180
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmBasePlate_Fixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsin_Cdepth As Single
Dim fsin_Cwidth As Single
Dim fsin_Cwt As Single
Dim fsin_Cft As Single
Dim fbl_lstColumn As Boolean
Dim gstr_Job As String
Dim gs_PMLunit As String

Private Sub chkI_Click()
If chkI.Value = 0 Then
      txtI.Enabled = False
      txtIfactor.Visible = True
      Label16.Visible = True
      txtI.BackColor = &H808080
Else
      txtI.Enabled = True
      txtIfactor.Visible = False
      Label16.Visible = False
      txtI.BackColor = &H80000005
End If

End Sub

Private Sub cmbCode_Click()

Dim lstr_Code As String
Dim xSQL As String

lstr_Code = CStr(Trim(cmbCode.Text))
    

If lstr_Code = "JIS" Then
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
               "where member_type = 'hbeam' " & _
               "order by member_no "
               cmbColumn.Enabled = True
Else
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
              "where member_type = 'hbeam' "
               cmbColumn.Enabled = False
End If
Call Query_AddList2_function(0, lstColumn, xSQL)


End Sub

Private Sub cmbColumn_Click()
Dim sql As String

If cmbColumn.Text <> "" Then
    sql = "Select * from code_jis " & _
          "where member_type = 'hbeam' " & _
          "and member_sort = '" & cmbColumn.Text & "' " & _
          "order by member_no "
Else
    sql = "Select * from code_jis " & _
          "where member_type = 'hbeam' " & _
          "order by member_no "
End If

Call Query_AddList_function(0, lstColumn, sql)
End Sub

Private Sub cmbMake_Click()
Dim TempPath As String, MemberName As String, TempBoltName As String
Dim TempUnit As String, TempMemUnit As String, TempDir As String, TempNut As Integer
Dim TempA As Single, TempB As Single, TempC As Single, TempD As Single, TempF As Single, TempG As Single, TempBPt As Single
Dim TempRBt As Single, TempRBr As Single, TempRBe As Single, TempRBh As Single
Dim lstr_Code As String

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

If Opt01.Value = True Then
    TempDir = "VectorY"
Else
    TempDir = "VectorX"
End If
TempNut = CInt(chkNut.Value)

lstr_Code = CStr(Trim(txtCode.Text))

MemberName = txtName.Text
TempBoltName = gfunSS_GetText(ssData, ssData.ActiveRow, 22)
TempMemUnit = cmbUnit.Text

TempA = CSng(Trim(txtA.Text))
TempB = CSng(Trim(txtB.Text))
TempC = CSng(Trim(txtC.Text))
TempD = CSng(Trim(txtD.Text))
TempF = CSng(Trim(txtF.Text))
TempG = CSng(Trim(txtG.Text))
TempBPt = CSng(Trim(txtBPT.Text))

TempRBt = CSng(Trim(txtRBT.Text))
TempRBr = CSng(Trim(txtJ.Text))
TempRBe = CSng(Trim(txtJ.Text))
TempRBh = CSng(Trim(txtI.Text))

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
    Call BasePlate_Fixed_PML(TempPath, gstr_Job, lstr_Code, MemberName, TempBoltName, _
                             gstr_BPF_Flag, TempUnit, TempMemUnit, TempDir, TempNut, _
                             TempA, TempB, TempC, TempD, TempF, TempG, TempBPt, _
                             TempRBt, TempRBr, TempRBe, TempRBh)
  Call PML_Run(TempPath)
  End
End If
Labelstop:

End Sub

Private Sub cmbUnit_Click()

If Trim(txtName.Text) = "" Then
      MsgBox "Select or Input the Name of Steel Member..."
Else
      Call Ribplate_Data(CStr(Trim(txtName.Text)), CStr(Trim(cmbCode.Text)))
End If

End Sub

Private Sub cmdDelete_Click()
Call datadelete
txtName.Text = ""
txtA.Text = "0"
txtB.Text = "0"
txtC.Text = "0"
txtD.Text = "0"
txtE.Text = "0"
txtF.Text = "0"
txtG.Text = "0"
cmbUnit.Text = "mm"

txtH.Text = "0"
txtIfactor.Text = "1.5"
txtI.Text = "0"
txtJ.Text = "0"
txtDM.Text = "A"
txtRDN.Text = "AC-15"
txtBPT.Text = "0"
txtRBT.Text = "0"
txtCode.Text = "0"
End Sub

Private Sub cmdExit_Click()
Dim i As Integer
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

Private Sub cmdSave_Click()

Call DataSave

End Sub

Private Sub Form_Load()

Dim xSQL As String
Dim i
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
      Frame3.Caption = "Current Project : " & gstr_Job & " , View Saved Data"
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
End If


fbl_lstColumn = True
'Me.Top = 0
'Me.Left = 0
Select Case gstr_BPF_Flag
    Case "Type01"
        Set imgFeature.Picture = ImageList1.ListImages(1).Picture
        Set Image3.Picture = ImageList1.ListImages(8).Picture
    Case "Type02"
        Set imgFeature.Picture = ImageList1.ListImages(2).Picture
        Set Image3.Picture = ImageList1.ListImages(8).Picture
    Case "Type03"
        Set imgFeature.Picture = ImageList1.ListImages(3).Picture
        Set Image3.Picture = ImageList1.ListImages(8).Picture
    Case "Type04"
        Set imgFeature.Picture = ImageList1.ListImages(4).Picture
        Set Image3.Picture = ImageList1.ListImages(8).Picture
    Case "Type05"
        Set imgFeature.Picture = ImageList1.ListImages(5).Picture
        Set Image3.Picture = ImageList1.ListImages(8).Picture
    Case "Type06"
        Set imgFeature.Picture = ImageList1.ListImages(6).Picture
        txtH.Enabled = False
        txtIfactor.Enabled = False
        txtJ.Enabled = False
    Case "Type07"
        Set imgFeature.Picture = ImageList1.ListImages(7).Picture
        txtH.Enabled = False
        txtIfactor.Enabled = False
        txtJ.Enabled = False
End Select

Set Image1.Picture = ImageList2.ListImages(1).Picture
Set Image2.Picture = ImageList2.ListImages(2).Picture

cmbUnit.AddItem "mm"
cmbUnit.AddItem "inch"

Call gs_CobAddItem(cmbCode)

'cmbCode.AddItem "JIS"
'cmbCode.AddItem "AISC"

Call InputText_Control

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
Else

               xSQL = "select Member_Name, A, B, C, D, E, F, G, Unit,  Type, H, Ifactor, I, J, Cdepth, " & _
                              "Cwidth, Cwt, Cft, DM, RDN, BPT, BoltName, RBT ,code from BasePlate_Fixed " & _
                              "where type = '" & gstr_BPF_Flag & "' and job ='" & gstr_Job & "'"

               Call gsubSSADOQuery(1, xSQL, ssData)
End If

'xSQL = "select member_name from code_JIS where member_type = 'hbeam' order by member_no"
'Call Query_AddList_function(lstColumn, xSQL)

xSQL = "Select member_sort from code_jis " & _
      "where member_type = 'hbeam' " & _
      "group by member_sort " & _
      "order by member_sort "

Call Query_AddList2_function(0, cmbColumn, xSQL)
Call cmbColumn_Click

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

Private Sub lstColumn_Click()

fbl_lstColumn = True
txtName.Text = Trim(lstColumn.List(lstColumn.ListIndex))

txtA.Text = "0"
txtB.Text = "0"
txtC.Text = "0"
txtD.Text = "0"
txtE.Text = "0"
txtF.Text = "0"
txtG.Text = "0"
'cmbUnit.Text = "mm"

txtH.Text = "0"
txtIfactor.Text = "1.5"
txtI.Text = "0"
txtJ.Text = "0"
txtDM.Text = "A"
txtRDN.Text = "AC-15"
txtBPT.Text = "0"
txtRBT.Text = "0"

End Sub
Private Sub txtB_Change()

Call cal_Hvalue

End Sub

Private Sub txtJ_Change()
Call cal_Hvalue

End Sub
Private Sub ssData_Click(ByVal Col As Long, ByVal Row As Long)

Dim xSQL As String
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, _
    Temp5 As String, Temp6 As String, Temp7 As String, Temp8 As String, _
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, _
    Temp13 As String, Temp14 As String, Temp15 As String, Tmep16 As String, _
    Temp17 As String, Temp18 As String, Temp19 As String, temp20 As String, _
    Temp21 As String, Temp22 As String, Temp23 As String, Temp24 As String
    
fbl_lstColumn = False

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
Temp17 = gfunSS_GetText(ssData, ssData.ActiveRow, 17)
Temp18 = gfunSS_GetText(ssData, ssData.ActiveRow, 18)
Temp19 = gfunSS_GetText(ssData, ssData.ActiveRow, 19)
temp20 = gfunSS_GetText(ssData, ssData.ActiveRow, 20)
Temp21 = gfunSS_GetText(ssData, ssData.ActiveRow, 21)
Temp22 = gfunSS_GetText(ssData, ssData.ActiveRow, 22)
Temp23 = gfunSS_GetText(ssData, ssData.ActiveRow, 23)
Temp24 = gfunSS_GetText(ssData, ssData.ActiveRow, 24)

gsin_Cdepth = convert(CSng(Trim(Temp15)), "mm", Temp9)
gsin_Cwidth = convert(CSng(Trim(Temp16)), "mm", Temp9)
gsin_Cwt = convert(CSng(Trim(Temp17)), "mm", Temp9)
gsin_Cft = convert(CSng(Trim(Temp18)), "mm", Temp9)

txtName.Text = Temp1
txtA.Text = Temp2
txtB.Text = Temp3
txtC.Text = Temp4
txtD.Text = Temp5
txtE.Text = Temp6
txtF.Text = Temp7
txtG.Text = Temp8
cmbUnit.Text = Temp9
gstr_BPFtype = CStr(Trim(Temp10))

'txtH.Text = Temp11
txtIfactor.Text = Temp12
txtJ.Text = Temp14

txtI.Text = Temp13

txtDM.Text = Temp19
txtRDN.Text = temp20
txtBPT.Text = Temp21
txtBname.Text = Temp22
txtRBT.Text = Temp23
txtCode.Text = Temp24
End Sub

Private Sub datadelete()

Dim xSQL As String
Dim Temp1 As String, Temp2 As String, Temp3 As String, Temp4 As String, _
    Temp5 As String, Temp6 As String, Temp7 As String, Temp8 As String, _
    Temp9 As String, Temp10 As String, Temp11 As String, Temp12 As String, _
    Temp13 As String, Temp14 As String, Temp15 As String, Tmep16 As String, _
    Temp17 As String, Temp18 As String, Temp19 As String, temp20 As String, _
    Temp21 As String, Temp22 As String, Temp23 As String, Temp24 As String

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
Temp17 = gfunSS_GetText(ssData, ssData.ActiveRow, 17)
Temp18 = gfunSS_GetText(ssData, ssData.ActiveRow, 18)
Temp19 = gfunSS_GetText(ssData, ssData.ActiveRow, 19)
temp20 = gfunSS_GetText(ssData, ssData.ActiveRow, 20)
Temp21 = gfunSS_GetText(ssData, ssData.ActiveRow, 21)
Temp22 = gfunSS_GetText(ssData, ssData.ActiveRow, 22)
Temp23 = gfunSS_GetText(ssData, ssData.ActiveRow, 23)
Temp24 = gfunSS_GetText(ssData, ssData.ActiveRow, 24)

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else

               xSQL = "delete from BasePlate_fixed where Member_Name = '" & CStr(Trim(Temp1)) & "' "
               xSQL = xSQL & "and A = " & CSng(Trim(Temp2)) & " "
               xSQL = xSQL & "and B = " & CSng(Trim(Temp3)) & " "
               xSQL = xSQL & "and C = " & CSng(Trim(Temp4)) & " "
               xSQL = xSQL & "and D = " & CSng(Trim(Temp5)) & " "
               xSQL = xSQL & "and E = " & CSng(Trim(Temp6)) & " "
               xSQL = xSQL & "and F = " & CSng(Trim(Temp7)) & " "
               xSQL = xSQL & "and G = " & CSng(Trim(Temp8)) & " "
               xSQL = xSQL & "and Unit = '" & CStr(Trim(Temp9)) & "' "
               xSQL = xSQL & "and Type = '" & CStr(Trim(Temp10)) & "' "
               xSQL = xSQL & "and H = " & CSng(Trim(Temp11)) & " "
               xSQL = xSQL & "and Ifactor = " & CSng(Trim(Temp12)) & " "
               xSQL = xSQL & "and I = " & CSng(Trim(Temp13)) & " "
               xSQL = xSQL & "and J = " & CSng(Trim(Temp14)) & " "
               xSQL = xSQL & "and Cdepth = " & CSng(Trim(Temp15)) & " "
               xSQL = xSQL & "and Cwidth = " & CSng(Trim(Temp16)) & " "
               xSQL = xSQL & "and Cwt = " & CSng(Trim(Temp17)) & " "
               xSQL = xSQL & "and Cft = " & CSng(Trim(Temp18)) & " "
               xSQL = xSQL & "and DM = '" & CStr(Trim(Temp19)) & "' "
               xSQL = xSQL & "and RDN = '" & CStr(Trim(temp20)) & "' "
               xSQL = xSQL & "and BPT = " & CSng(Trim(Temp21)) & " "
               xSQL = xSQL & "and BoltName = '" & CStr(Trim(Temp22)) & "' "
               xSQL = xSQL & "and RBT = " & CSng(Trim(Temp23)) & " "
               xSQL = xSQL & "and job = '" & gstr_Job & "'"
               xSQL = xSQL & "and Code = '" & CStr(Trim(Temp24)) & "'"
               
               adoConnection1.Execute (xSQL)
                     
                     
               If ssData.ActiveRow = 1 Then Row = 1
               If ssData.ActiveRow = ssData.MaxRows Then Row = ssData.MaxRows - 1
               
               Call gsubSS_DelRow(ssData, ssData.ActiveRow)
               'If ssData.ActiveCol = ssData.MaxCols Then ssData
               xSQL = "select Member_Name, A, B, C, D, E, F, G, Unit,  Type, H, Ifactor, I, J, Cdepth, " & _
                              "Cwidth, Cwt, Cft, DM, RDN, BPT, BoltName, RBT from BasePlate_Fixed " & _
                              "where type = '" & gstr_BPF_Flag & "' and job ='" & gstr_Job & "'"
                              
               Call gsubSSADOQuery(1, xSQL, ssData)
End If
End Sub

Private Sub InputText_Control()

Select Case gstr_BPF_Flag
    Case "Type01"
        txtA.Enabled = True
        txtB.Enabled = True
        txtC.Enabled = True
        txtD.Enabled = True
        txtE.Enabled = True
        txtF.Enabled = True
        txtG.Enabled = False
    Case "Type02"
        txtA.Enabled = True
        txtB.Enabled = True
        txtC.Enabled = True
        txtD.Enabled = True
        txtE.Enabled = True
        txtF.Enabled = True
        txtG.Enabled = True
    Case "Type03", "Type04", "Type05", "Type06", "Type07"
        txtA.Enabled = True
        txtB.Enabled = True
        txtC.Enabled = True
        txtD.Enabled = True
        txtE.Enabled = True
        txtF.Enabled = False
        txtG.Enabled = False
End Select


End Sub
Private Sub DataSave()

Dim xSQL As String
If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General Window"
               Exit Sub
Else
               xSQL = "select Member_Name from BasePlate_Fixed where Member_Name = '" & CStr(Trim(txtName.Text)) & "' "
               xSQL = xSQL & "and type = '" & gstr_BPF_Flag & "' "
               xSQL = xSQL & " and job = '" & gstr_Job & "'"
               
               If DataCheck(CStr(Trim(txtName.Text)), xSQL) Then
                   xSQL = "update BasePlate_Fixed set "
                   xSQL = xSQL & "Member_Name ='" & CStr(Trim(txtName.Text)) & "', "
                   xSQL = xSQL & "A = " & CSng(Trim(txtA.Text)) & ","
                   xSQL = xSQL & "B = " & CSng(Trim(txtB.Text)) & ","
                   xSQL = xSQL & "C = " & CSng(Trim(txtC.Text)) & ","
                   xSQL = xSQL & "D = " & CSng(Trim(txtD.Text)) & ", "
                   xSQL = xSQL & "E = " & CSng(Trim(txtE.Text)) & ", "
                   xSQL = xSQL & "F = " & CSng(Trim(txtF.Text)) & ", "
                   xSQL = xSQL & "G = " & CSng(Trim(txtG.Text)) & ", "
                   xSQL = xSQL & "Unit = '" & CStr(Trim(cmbUnit.Text)) & "', "
                   xSQL = xSQL & "Type ='" & gstr_BPF_Flag & "', "
                   xSQL = xSQL & "H = " & CSng(Trim(txtH.Text)) & ", "
                   xSQL = xSQL & "Ifactor = " & CSng(Trim(txtIfactor.Text)) & ", "
                   xSQL = xSQL & "I = " & CSng(Trim(txtI.Text)) & ", "
                   xSQL = xSQL & "J = " & CSng(Trim(txtJ.Text)) & ", "
                   xSQL = xSQL & "Cdepth = " & gsin_Cdepth & ", "
                   xSQL = xSQL & "Cwidth = " & gsin_Cwidth & ", "
                   xSQL = xSQL & "Cwt = " & gsin_Cwt & ", "
                   xSQL = xSQL & "Cft = " & gsin_Cft & ", "
                   xSQL = xSQL & "DM = '" & CStr(Trim(txtDM.Text)) & "', "
                   xSQL = xSQL & "RDN = '" & CStr(Trim(txtRDN.Text)) & "', "
                   xSQL = xSQL & "BPT = " & CSng(Trim(txtBPT.Text)) & ", "
                   xSQL = xSQL & "BoltName = '" & Trim(txtBname.Text) & "', "
                   xSQL = xSQL & "RBT = " & CSng(Trim(txtRBT.Text)) & " "
                   xSQL = xSQL & "Where Member_Name = '" & CStr(Trim(txtName.Text)) & "'"
                   xSQL = xSQL & " and type = '" & gstr_BPF_Flag & "' "
                   xSQL = xSQL & "and job = '" & gstr_Job & "'"
                   adoConnection1.Execute (xSQL)
                   
               Else
                   If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
                   xSQL = "insert into BasePlate_Fixed values ('"
                   xSQL = xSQL & gstr_Job & "', '"
                   xSQL = xSQL & CStr(Trim(txtName.Text)) & "',"
                   xSQL = xSQL & CSng(Trim(txtA.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtB.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtC.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtD.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtE.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtF.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtG.Text)) & ", '"
                   xSQL = xSQL & CStr(Trim(cmbUnit.Text)) & "', '"
                   xSQL = xSQL & gstr_BPF_Flag & "',"
                   xSQL = xSQL & CSng(Trim(txtH.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtIfactor.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtI.Text)) & ","
                   xSQL = xSQL & CSng(Trim(txtJ.Text)) & ","
                   xSQL = xSQL & fsin_Cdepth & ","
                   xSQL = xSQL & fsin_Cwidth & ","
                   xSQL = xSQL & fsin_Cwt & ","
                   xSQL = xSQL & fsin_Cft & ", '"
                   xSQL = xSQL & CStr(Trim(txtDM.Text)) & "', '"
                   xSQL = xSQL & CStr(Trim(txtRDN.Text)) & "', "
                   xSQL = xSQL & CSng(Trim(txtBPT.Text)) & ", '"
                   xSQL = xSQL & Trim(txtBname.Text) & "', "
                   xSQL = xSQL & CSng(Trim(txtRBT.Text)) & ", '"
                   xSQL = xSQL & CStr(Trim(cmbCode.Text)) & "')"
                   adoConnection1.Execute (xSQL)
               
               End If
               xSQL = "select Member_Name, A, B, C, D, E, F, G, Unit,  Type, H, Ifactor, I, J, Cdepth, " & _
                              "Cwidth, Cwt, Cft, DM, RDN, BPT, BoltName, RBT , Code from BasePlate_Fixed " & _
                              "where type = '" & gstr_BPF_Flag & "' and job ='" & gstr_Job & "'"
               
               Call gsubSSADOQuery(1, xSQL, ssData)
End If

End Sub

Private Sub Ribplate_Data(valName As String, ByVal valCode As String)
Dim retData As ADODB.Recordset
Dim xSQL As String
xSQL = "select * from code_" & valCode & " where member_name = '" & valName & "'"

Set retData = adoConnection.Execute(xSQL)
    fsin_Cdepth = retData!D
    fsin_Cwidth = retData!Bf
    fsin_Cwt = retData!Tw
    fsin_Cft = retData!Tf
retData.Close

Set retData = Nothing

fsin_Cdepth = convert(fsin_Cdepth, "mm", CStr(Trim(cmbUnit.Text)))
fsin_Cwidth = convert(fsin_Cwidth, "mm", CStr(Trim(cmbUnit.Text)))
fsin_Cwt = convert(fsin_Cwt, "mm", CStr(Trim(cmbUnit.Text)))
fsin_Cft = convert(fsin_Cft, "mm", CStr(Trim(cmbUnit.Text)))


txtH.Text = (CSng(Trim(txtB.Text)) - fsin_Cdepth) / 2
txtI.Text = CSng(Trim(txtH.Text)) * CSng(Trim(txtIfactor.Text))
End Sub



Private Sub cal_Hvalue()
Dim TempB As Single, TempC As Single

If CStr(Trim(txtB.Text)) = "" Then
    TempB = 0
Else
    TempB = CSng(Trim(txtB.Text))
End If

If CStr(Trim(txtJ.Text)) = "" Then
    TempC = 0
Else
    TempC = CSng(Trim(txtJ.Text))
End If

If TempB = 0 Or TempC = 0 Then
    GoTo label
Else
    If fbl_lstColumn = True Then
        txtH.Text = (TempB - fsin_Cdepth) / 2 - TempC
    Else
        txtH.Text = (TempB - gsin_Cdepth) / 2 - TempC
    End If
    txtI.Text = CSng(Trim(txtH.Text)) * CSng(Trim(txtIfactor.Text))

End If

label:
End Sub


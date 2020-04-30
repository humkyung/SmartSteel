VERSION 5.00
Begin VB.Form frmMoudle_Hor 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmModule_Hor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10050
   Begin VB.Frame Frame6 
      Height          =   765
      Left            =   6720
      TabIndex        =   34
      Top             =   4770
      Width           =   3255
      Begin VB.CheckBox chkNut 
         Caption         =   "Nut Model Check"
         Height          =   315
         Left            =   600
         TabIndex        =   35
         Top             =   270
         Value           =   1  '확인
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selcet Member Code"
      Height          =   645
      Left            =   6720
      TabIndex        =   30
      Top             =   30
      Width           =   3255
      Begin VB.ComboBox cmbCode 
         Height          =   300
         Left            =   1170
         TabIndex        =   31
         Text            =   "JIS"
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmbMake 
      Caption         =   "Make PML"
      Height          =   480
      Left            =   6840
      TabIndex        =   29
      Top             =   6600
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select PML Unit"
      Height          =   930
      Left            =   60
      TabIndex        =   22
      Top             =   6150
      Width           =   3255
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   390
         Width           =   735
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   885
         TabIndex        =   25
         Top             =   420
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   1500
         TabIndex        =   24
         Top             =   420
         Width           =   690
      End
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   2355
         TabIndex        =   23
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Select Modeling Option"
      Height          =   1320
      Left            =   30
      TabIndex        =   18
      Top             =   4770
      Width           =   3300
      Begin VB.ComboBox cmbSubType 
         Height          =   300
         ItemData        =   "frmModule_Hor.frx":000C
         Left            =   270
         List            =   "frmModule_Hor.frx":0019
         TabIndex        =   21
         Text            =   "Type-a"
         Top             =   870
         Width           =   2715
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   270
         TabIndex        =   20
         Text            =   "Type-01"
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type of Module-01 :"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   285
         Width           =   1725
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "View of Modeling Option"
      Height          =   2340
      Left            =   3375
      TabIndex        =   17
      Top             =   4770
      Width           =   3300
      Begin VB.TextBox txtSpace 
         Height          =   270
         Left            =   2190
         TabIndex        =   33
         Text            =   "0"
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblSpace 
         AutoSize        =   -1  'True
         Caption         =   "Space"
         Height          =   180
         Left            =   2190
         TabIndex        =   32
         Top             =   990
         Width           =   540
      End
      Begin VB.Image imgModel 
         Height          =   2100
         Left            =   45
         Picture         =   "frmModule_Hor.frx":0035
         Top             =   180
         Width           =   3000
      End
   End
   Begin VB.Frame fraModule 
      Caption         =   "View of Module-01"
      Height          =   2310
      Left            =   3375
      TabIndex        =   16
      Top             =   45
      Width           =   3300
      Begin VB.Image imgModule 
         Height          =   2025
         Left            =   315
         Picture         =   "frmModule_Hor.frx":14897
         Top             =   225
         Width           =   2850
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "View Axis"
      Height          =   2310
      Left            =   45
      TabIndex        =   15
      Top             =   45
      Width           =   3300
      Begin VB.Image imgAxis 
         Height          =   1905
         Left            =   225
         Picture         =   "frmModule_Hor.frx":2767D
         Top             =   270
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Your Selection"
      Height          =   1590
      Left            =   6705
      TabIndex        =   8
      Top             =   765
      Width           =   3300
      Begin VB.Label lblBracing 
         AutoSize        =   -1  'True
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   225
         TabIndex        =   14
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblSubBeam 
         AutoSize        =   -1  'True
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   885
         Width           =   345
      End
      Begin VB.Label lblBeam 
         AutoSize        =   -1  'True
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   450
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bracing Size :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1140
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Beam Size   :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   705
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Beam Size :"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   8505
      TabIndex        =   6
      Top             =   6615
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Bracing"
      Height          =   2310
      Left            =   6705
      TabIndex        =   4
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstBType 
         Height          =   780
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   3165
      End
      Begin VB.ListBox lstBracing 
         Columns         =   2
         Height          =   1140
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   3165
      End
   End
   Begin VB.Frame fraSubBeam 
      Caption         =   "Select Sub Beam"
      Height          =   2310
      Left            =   3375
      TabIndex        =   1
      Top             =   2430
      Width           =   3300
      Begin VB.ComboBox cmbSubBeam 
         Height          =   300
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   3165
      End
      Begin VB.ListBox lstSubBeam 
         Columns         =   2
         Height          =   1500
         Left            =   90
         TabIndex        =   3
         Top             =   630
         Width           =   3165
      End
   End
   Begin VB.Frame fraBeam 
      Caption         =   "Select Beam"
      Height          =   2310
      Left            =   45
      TabIndex        =   0
      Top             =   2430
      Width           =   3300
      Begin VB.ComboBox cmbBeam 
         Height          =   300
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   3165
      End
      Begin VB.ListBox lstBeam 
         Columns         =   2
         Height          =   1500
         Left            =   60
         TabIndex        =   2
         Top             =   630
         Width           =   3165
      End
   End
End
Attribute VB_Name = "frmMoudle_Hor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spec_Code As String
Dim gstr_Job As String, gs_PMLunit As String

Private Sub cmbColShape_click()

Select Case gstr_VBGP_Flag
   Case "Module01", "Module02"
    If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
    Else
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
    End If
   Case "Module03", "Module04"
    If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
    Else
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
    End If
End Select

'Call Type_BMP_Control_Start
End Sub

Private Sub cmbBeam_Click()
Dim sql As String

If cmbBeam.Text <> "" Then
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "and member_sort = '" & cmbBeam.Text & "' " & _
          "order by member_no "
Else
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "order by member_no "
End If

Call Query_AddList_function(0, lstBeam, sql)
End Sub

Private Sub cmbCode_Click()

Dim formCode As String
Dim xSQL As String

formCode = CStr(Trim(cmbCode.Text))
    

If formCode = "JIS" Then
      
      cmbBeam.Enabled = True
      cmbSubBeam.Enabled = True
      
      xSQL = "Select member_sort from code_" & formCode & " " & _
      "where member_type = 'Hbeam' " & _
      "group by member_sort " & _
      "order by member_sort "
      
      Call Query_AddList2_function(0, cmbSubBeam, xSQL)
      'Call cmbSubBeam_Click
      Call Query_AddList2_function(0, cmbBeam, xSQL)
      'Call cmbBeam_Click
      
      xSQL = "Select member_name from code_" & formCode & " " & _
      "where member_type = 'hbeam' " & _
      "order by member_no "
Else
       xSQL = "Select member_name from code_" & formCode & " " & _
      "where member_type = 'hbeam' "
       cmbBeam.Enabled = False
       cmbSubBeam.Enabled = False
               
End If
Call Query_AddList2_function(0, lstBeam, xSQL)
Call Query_AddList2_function(0, lstSubBeam, xSQL)



End Sub

Private Sub cmbSubBeam_Click()
Dim sql As String

If cmbSubBeam.Text <> "" Then
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "and member_sort = '" & cmbSubBeam.Text & "' " & _
          "order by member_no "
Else
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "order by member_no "
End If

Call Query_AddList_function(0, lstSubBeam, sql)

End Sub

Private Sub cmbMake_Click()
Dim SubBeamName As String
Dim BeamName As String
Dim BracingName As String
Dim tempColShape As String, tempType As String, TempUnit As String, TempPath As String
Dim tempSpace3 As Single, TempNutFlag As Integer
Dim formCode As String

On Error GoTo Labelstop


    SubBeamName = lblSubBeam.Caption
    BeamName = lblBeam.Caption
    BracingName = lblBracing.Caption
    tempSpace3 = CSng(Trim(txtSpace.Text))
    TempNutFlag = CInt(Trim(chkNut.Value))
    formCode = CStr(Trim(cmbCode.Text))
    If formCode = "" Then MsgBox "Code Selection Error. You must select Code Selection !!!": Exit Sub

    Select Case gstr_HBGP_Flag
        Case "Module01", "Module02", "Module03", "Module04", "Module06"
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
        Case "Module07"
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
            If SubBeamName = "N/A" Then
                MsgBox "Sub Beam을 선택 하십시요."
                Exit Sub
            End If
         Case "Module05"
               If tempSpace3 <= 0 Then
                              MsgBox "Space Value is Zero or Minus. Chage Space Value"
                Exit Sub
            End If
    End Select
    If BracingName = "N/A" Then
        MsgBox "Bracing을 선택 하십시요."
        Exit Sub
    End If
    
    tempType = cmbType.Text

    If optMM.Value = True Then
        TempUnit = "mm"
    ElseIf optM.Value = True Then
        TempUnit = "m"
    ElseIf optInch.Value = True Then
        TempUnit = "inch"
    ElseIf optFeet.Value = True Then
        TempUnit = "feet"
    End If

      Open App.Path & "\Library\" & gstr_HBGP_Flag & "_Hori_history.ini" For Output As #1
            Print #1, gstr_Job
            Print #1, Trim(cmbCode.Text)
            Print #1, Trim(lstBeam.ListIndex)
            Print #1, Trim(lstSubBeam.ListIndex)
            Print #1, Trim(lstBType.ListIndex)
            Print #1, Trim(lstBracing.ListIndex)
      Close #1

'    If BeamName = "" Then
'        MsgBox "Select Beam Member Size....."
'    Else
'        frmMain.CommonDialog.CancelError = True
        frmMain.CommonDialog.InitDir = App.Path
        frmMain.CommonDialog.DialogTitle = "Save PML File "
        frmMain.CommonDialog.Filter = "BasePlate (*.pml)|*.pml|"
        frmMain.CommonDialog.FileName = "Test.pml"
        
        frmMain.CommonDialog.ShowSave
        TempPath = frmMain.CommonDialog.FileName
    
        Call HB_PML(TempPath, gstr_Job, spec_Code, formCode, gstr_HBGP_Flag, tempType, TempUnit, _
                    BeamName, SubBeamName, BracingName, tempSpace3, TempNutFlag)
        Call PML_Run(TempPath)
        End
'    End If

Labelstop:
End Sub

Private Sub cmbSubType_Click()
Call SubType_BMP_Control
End Sub

Private Sub cmbType_Click()
Call Type_BMP_Control
End Sub

Private Sub cmdExit_Click()

If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub Form_Load()

Dim sql As String
Dim i
'Call gs_project_Input(gstr_Job)


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
      Me.Caption = "Current Project : " & gstr_Job & ", " & gs_Caption
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
End If


lstBType.AddItem "Angle"
lstBType.AddItem "Channel"
lstBType.AddItem "Double Angle"
lstBType.AddItem "Double Channel"
lstBType.AddItem "Tee"

Call gs_CobAddItem(cmbCode)

Dim pre_Code As String, pre_BeamIndex As String, pre_subBeamIndex As String, pre_BtypeIndex As String, pre_BraceIndex As String, _
    pre_Job As String
On Error GoTo Error100
Open App.Path & "\Library\" & gstr_HBGP_Flag & "_Hori_history.ini" For Input As #1
      Input #1, pre_Job
      Input #1, pre_Code
      Input #1, pre_BeamIndex
      Input #1, pre_subBeamIndex
      Input #1, pre_BtypeIndex
      Input #1, pre_BraceIndex
Close #1

restart01:

cmbCode.Text = pre_Code
Call cmbCode_Click
If pre_Code = "JIS" Then
      sql = "Select member_sort from code_" & pre_Code & " " & _
            "where member_type = 'Hbeam' " & _
            "group by member_sort " & _
            "order by member_sort "

      Call Query_AddList2_function(0, cmbSubBeam, sql)
      Call cmbSubBeam_Click
      Call Query_AddList2_function(0, cmbBeam, sql)
      Call cmbBeam_Click
End If

If gstr_Job = pre_Job Then
      If pre_BeamIndex <> "-1" And pre_BeamIndex <> "" Then
            lstBeam.Selected(pre_BeamIndex) = True
      End If
      If pre_subBeamIndex <> "-1" And pre_subBeamIndex <> "" Then
            lstSubBeam.Selected(pre_subBeamIndex) = True
      End If
      If pre_BtypeIndex <> "-1" And pre_BtypeIndex <> "" Then
            lstBType.Selected(pre_BtypeIndex) = True
            
      End If
      
      If pre_BraceIndex <> "-1" And pre_BraceIndex <> "" Then
            lstBracing.Selected(pre_BraceIndex) = True
      End If
End If


If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
'Else
'               sql = "Select Member_Name from HB_Connection" & _
'                     " where Shape = 'Angle' and job = '" & gstr_Job & "'"
'
'               Call Query_AddList_function(1, lstBracing, sql)
End If


Call Form_Control

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
Exit Sub

Error100:
Select Case Err.Number
   Case 53   ' "txt파일이 없는 경우의 error 번호(path에 따라 다른 것 같음)
      pre_Code = "JIS"
      GoTo restart01
   Case Else
      ' 여기서 다른 상황을 다룹니다.
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub


Private Sub lstBType_Click()
Dim strTemp As String
Dim sql As String


If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else
               
               strTemp = Trim(lstBType.List(lstBType.ListIndex))
               
               sql = "Select Member_Name from HB_Connection" & _
                         " where Shape = '" & strTemp & "' and job = '" & gstr_Job & "'"
               
               Call Query_AddList_function(1, lstBracing, sql)
End If
'opt01.Enabled = True
'opt02.Enabled = True
'opt03.Enabled = True
'opt04.Enabled = True
'opt05.Enabled = True
'opt06.Enabled = True

'Select Case strTemp
'    Case "Angle"
'        opt01.Value = True
'        opt03.Enabled = False
'        opt04.Enabled = False
'        opt05.Enabled = False
'        opt06.Enabled = False
'    Case "Channel"
'        opt01.Enabled = False
'        opt02.Enabled = False
'        opt03.Value = True
''        opt04.Enabled = False
'        opt05.Enabled = False
'        opt06.Enabled = False
'    Case "Double Angle"
'        opt01.Value = True
'        opt02.Enabled = False
'        opt03.Enabled = False
'        opt04.Enabled = False
''        opt05.Enabled = False
''        opt06.Enabled = False
'    Case "Double Channel"
'        opt01.Enabled = False
'        opt02.Enabled = False
'        opt03.Value = True
'        opt04.Enabled = False
'        opt05.Enabled = False
'        opt06.Enabled = False
'
'    Case "Tee"
'        opt06.Value = True
'        opt02.Enabled = False
'        opt03.Enabled = False
'        opt04.Enabled = False
''        opt05.Enabled = False
''        opt06.Enabled = False

'End Select

End Sub

Private Sub lstSubBeam_Click()

lblSubBeam.Caption = Trim(lstSubBeam.List(lstSubBeam.ListIndex))

End Sub
Private Sub lstBeam_Click()
lblBeam.Caption = Trim(lstBeam.List(lstBeam.ListIndex))

End Sub
Private Sub lstBracing_Click()
Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset

lblBracing.Caption = Trim(lstBracing.List(lstBracing.ListIndex))
lstr_MemberName = CStr(Trim(lstBracing.List(lstBracing.ListIndex)))

xSQL = "Select Code from HB_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

spec_Code = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing

End Sub

Private Sub Form_Control()

Select Case gstr_HBGP_Flag
    Case "Module01"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module01.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
        fraModule.Caption = "View of Module-01"
    Case "Module02"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module02.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
        fraModule.Caption = "View of Module-02"
    Case "Module03"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module03.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
        fraModule.Caption = "View of Module-03"
    Case "Module04"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module04.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
        fraModule.Caption = "View of Module-04"
    Case "Module05"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module05.bmp")
        fraBeam.Enabled = False
        lstBeam.Visible = False
        cmbBeam.Visible = False
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = True
        lblSpace.Caption = "Space(m):"
        txtSpace.Visible = True
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
'        lblColShape.Visible = False
'        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-05"
    Case "Module06"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module06.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
        cmbSubBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.ListIndex = 0
        cmbSubType.Visible = False
'        lblColShape.Visible = False
'        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-06"
    Case "Module07"
        imgModule.Picture = LoadPicture(App.Path & "\BMP\HB\Module\XY_Module07.bmp")
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = True
        lstSubBeam.Visible = True
        cmbSubBeam.Visible = True
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.ListIndex = 0
        cmbSubType.Visible = True
'        lblColShape.Visible = False
'        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-07"
End Select


Call Type_BMP_Control_Start


End Sub

Private Sub Type_BMP_Control_Start()

'If gstr_VBdir_Flag = "Y" Then
'    imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\YZ_Axis.bmp")
'Else
'    imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\XZ_Axis.bmp")
'End If



'Select Case gstr_VBGP_Flag
'    Case "Module01"
'        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type01_S.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type01_W.bmp")
'        End If
'    Case "Module02"
'        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_S.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_W.bmp")
'        End If
'    Case "Module03"
'        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_S.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_W.bmp")
'        End If
'
'    Case "Module04"
'        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_S.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_W.bmp")
'        End If
'
'    Case "Module05"
'        If CStr(Trim(cmbType.Text)) = "Type-01" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type01.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type02.bmp")
'        End If
'
'    Case "Module06"
'        If CStr(Trim(cmbType.Text)) = "Type-01" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type01.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type02.bmp")
'        End If
'
'    Case "Module07"
'        If CStr(Trim(cmbType.Text)) = "Type-01" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type01.bmp")
'        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type02.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type03.bmp")
'        End If
'
'    Case "Module08"
'        If CStr(Trim(cmbType.Text)) = "Type-01" Then
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type01.bmp")
'        Else
'            imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type02.bmp")
'        End If
'
'End Select

End Sub
Private Sub Type_BMP_Control()
Select Case gstr_HBGP_Flag
    Case "Module01"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type03.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-04" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type04.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-05" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type05.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-06" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module01_Type06.bmp")
        End If
    Case "Module02"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type03.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-04" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type04.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-05" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type05.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-06" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module02_Type06.bmp")
        End If
    Case "Module03"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type03.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-04" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type04.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-05" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type05.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-06" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module03_Type06.bmp")
        End If
    Case "Module04"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type03.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-04" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type04.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-05" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type05.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-06" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module04_Type06.bmp")
        End If
        
    Case "Module05"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module05_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module05_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module05_Type03.bmp")
        End If
       
    Case "Module06"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module06_Type01.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module06_Type02.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module06_Type03.bmp")
        ElseIf CStr(Trim(cmbType.Text)) = "Type-04" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module06_Type04.bmp")
        End If
        
    Case "Module07"
        cmbSubType.ListIndex = 0
        Call cmbSubType_Click
End Select

End Sub

Private Sub SubType_BMP_Control()
Select Case CStr(Trim(cmbType.Text))
    Case "Type-01"
        If CStr(Trim(cmbSubType.Text)) = "Type-a" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type01_a.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-b" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type01_b.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-c" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type01_c.bmp")
        End If
    Case "Type-02"
        If CStr(Trim(cmbSubType.Text)) = "Type-a" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type02_a.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-b" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type02_b.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-c" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type02_c.bmp")
        End If
    Case "Type-03"
        If CStr(Trim(cmbSubType.Text)) = "Type-a" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type03_a.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-b" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type03_b.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-c" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type03_c.bmp")
        End If
    Case "Type-04"
        If CStr(Trim(cmbSubType.Text)) = "Type-a" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type04_a.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-b" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type04_b.bmp")
        ElseIf CStr(Trim(cmbSubType.Text)) = "Type-c" Then
            imgModel.Picture = LoadPicture(App.Path & "\BMP\HB\Type\Module07_Type04_c.bmp")
        End If
End Select

End Sub

Private Sub optFeet_Click()
lblSpace.Caption = "Space(ft):"

End Sub

Private Sub optInch_Click()
lblSpace.Caption = "Space(in):"

End Sub

Private Sub optM_Click()
lblSpace.Caption = "Space(m):"
End Sub

Private Sub optMM_Click()
lblSpace.Caption = "Space(mm):"
End Sub

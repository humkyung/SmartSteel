VERSION 5.00
Begin VB.Form frmMoudle_Mc 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmModule_Mc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10050
   Begin VB.Frame Frame2 
      Caption         =   "Selcet Code"
      Height          =   615
      Left            =   6690
      TabIndex        =   23
      Top             =   60
      Width           =   3255
      Begin VB.ComboBox cmbCode 
         Height          =   300
         Left            =   1110
         TabIndex        =   24
         Text            =   "JIS"
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmbMake 
      Caption         =   "Make PML"
      Height          =   480
      Left            =   6840
      TabIndex        =   22
      Top             =   5400
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select PML Unit"
      Height          =   1140
      Left            =   60
      TabIndex        =   15
      Top             =   4770
      Width           =   3255
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   600
         TabIndex        =   19
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   1635
         TabIndex        =   18
         Top             =   330
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   600
         TabIndex        =   17
         Top             =   690
         Width           =   690
      End
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   1635
         TabIndex        =   16
         Top             =   690
         Width           =   600
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "View of Modeling Option"
      Height          =   2340
      Left            =   3375
      TabIndex        =   14
      Top             =   30
      Width           =   3300
      Begin VB.Image imgModel 
         Height          =   2100
         Left            =   105
         Picture         =   "frmModule_Mc.frx":000C
         Top             =   180
         Width           =   3000
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "View Axis"
      Height          =   2310
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   3300
      Begin VB.Image imgAxis 
         Height          =   1905
         Left            =   225
         Picture         =   "frmModule_Mc.frx":1486E
         Top             =   270
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Your Selection"
      Height          =   1590
      Left            =   6705
      TabIndex        =   6
      Top             =   765
      Width           =   3300
      Begin VB.Label lblRightBeam 
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
         Left            =   255
         TabIndex        =   12
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label lblLeftBeam 
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
         Left            =   255
         TabIndex        =   11
         Top             =   825
         Width           =   345
      End
      Begin VB.Label lblColumn 
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
         Left            =   255
         TabIndex        =   10
         Top             =   390
         Width           =   345
      End
      Begin VB.Label lblRightBeamTitle 
         AutoSize        =   -1  'True
         Caption         =   "Right Beam Size :"
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
         Top             =   1080
         Width           =   1905
      End
      Begin VB.Label lblLeftBeamTitle 
         AutoSize        =   -1  'True
         Caption         =   "Left Beam Size   :"
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
         TabIndex        =   8
         Top             =   660
         Width           =   1890
      End
      Begin VB.Label lblColumnTitle 
         AutoSize        =   -1  'True
         Caption         =   "Column Size :"
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
         TabIndex        =   7
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   8520
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame fraRightBeam 
      Caption         =   "Select Right Beam"
      Height          =   2310
      Left            =   6705
      TabIndex        =   3
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstRightBeam 
         Columns         =   2
         Height          =   1860
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.Frame fraLeftBeam 
      Caption         =   "Select Left Beam"
      Height          =   2310
      Left            =   3375
      TabIndex        =   1
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstLeftBeam 
         Columns         =   2
         Height          =   1860
         Left            =   60
         TabIndex        =   2
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.Frame fraColumn 
      Caption         =   "Select Column"
      Height          =   2310
      Left            =   45
      TabIndex        =   0
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstColumn 
         Columns         =   2
         Height          =   1500
         Left            =   60
         TabIndex        =   21
         Top             =   660
         Width           =   3165
      End
      Begin VB.ComboBox cmbColumn 
         Height          =   300
         Left            =   60
         TabIndex        =   20
         Top             =   300
         Width           =   3165
      End
   End
End
Attribute VB_Name = "frmMoudle_Mc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstr_Code_Left As String, lstr_Code_Right As String
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

Private Sub cmbSubType_Click()
Call SubType_BMP_Control
End Sub

Private Sub cmbType_Click()
Call Type_BMP_Control
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
'Call Query_AddList2_function(0, lstSubBeam, xSQL)



End Sub

Private Sub cmbColumn_Click()
Dim sql As String

If cmbColumn.Text <> "" Then
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "and member_sort = '" & cmbColumn.Text & "' " & _
          "order by member_no "
Else
    sql = "Select * from code_jis " & _
          "where member_type = 'Hbeam' " & _
          "order by member_no "
End If

Call Query_AddList_function(0, lstColumn, sql)

End Sub

Private Sub cmbMake_Click()
Dim TempPath As String, MemberName As String, TempBoltName As String
Dim TempUnit As String, TempMemUnit As String, TempDir As String, TempNut As Integer
Dim TempA As Single, TempB As Single, TempC As Single, TempD As Single, TempF As Single, TempG As Single, TempBPt As Single
Dim TempRBt As Single, TempRBr As Single, TempRBe As Single, TempRBh As Single
Dim ColumnName As String, LeftBeamName As String, RightBeamName As String
Dim formCode As String
On Error GoTo Labelstop


ColumnName = lblColumn.Caption
LeftBeamName = lblLeftBeam.Caption
RightBeamName = lblRightBeam.Caption
formCode = CStr(Trim(cmbCode.Text))
If formCode = "" Then MsgBox "Code Selection Error. You must select Code Selection !!!": Exit Sub

Select Case gstr_MCEP_type
    Case "Type01"
        If ColumnName = "N/A" Then
            MsgBox "Column을 선택 하십시요."
            Exit Sub
        End If
        If RightBeamName = "N/A" Then
            MsgBox "Beam을 선택 하십시요."
            Exit Sub
        End If
    Case "Type02"
        If ColumnName = "N/A" Then
            MsgBox "Column을 선택 하십시요."
            Exit Sub
        End If
        If LeftBeamName = "N/A" Then
            MsgBox "Beam을 선택 하십시요."
            Exit Sub
        End If
    Case "Type05"
        If ColumnName = "N/A" Then
            MsgBox "Column을 선택 하십시요."
            Exit Sub
        End If
        If LeftBeamName = "N/A" Then
            MsgBox "Beam을 선택 하십시요."
            Exit Sub
        End If
        RightBeamName = LeftBeamName
    Case Else
        If ColumnName = "N/A" Then
            MsgBox "Column을 선택 하십시요."
            Exit Sub
        End If
        If LeftBeamName = "N/A" Then
            MsgBox "Left Beam을 선택 하십시요."
            Exit Sub
        End If
        If RightBeamName = "N/A" Then
            MsgBox "Right Beam을 선택 하십시요."
            Exit Sub
        End If
End Select

If optMM.Value = True Then
    TempUnit = "mm"
ElseIf optM.Value = True Then
    TempUnit = "m"
ElseIf optInch.Value = True Then
    TempUnit = "inch"
ElseIf optFeet.Value = True Then
    TempUnit = "feet"
End If
'TempNut = CInt(chkNut.Value)
'
'MemberName = txtName.Text
'
'If MemberName = "" Then
'    MsgBox "Select Column Member Size....."
'Else

    frmMain.CommonDialog.CancelError = True
    frmMain.CommonDialog.InitDir = App.Path
    frmMain.CommonDialog.DialogTitle = "Save PML File "
    frmMain.CommonDialog.Filter = "BasePlate (*.pml)|*.pml|"
    frmMain.CommonDialog.FileName = "Test.pml"
    
    frmMain.CommonDialog.ShowSave
    TempPath = frmMain.CommonDialog.FileName
    
    If gstr_MCEP_type = "Type05" Then
            lstr_Code_Right = lstr_Code_Left
    End If
    
    Call MC_PML(TempPath, gstr_Job, lstr_Code_Left, lstr_Code_Right, formCode, gstr_MCdir_Flag, _
                                    gstr_MCEP_Flag, gstr_MCEP_type, TempUnit, _
                                    ColumnName, LeftBeamName, RightBeamName)
    Call PML_Run(TempPath)
    
    End
    
'End If
Labelstop:

End Sub

Private Sub cmdExit_Click()

If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub Form_Load()

Dim sql As String, WhereSql As String

'Me.Top = 0: Me.Left = 0
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
      fraType.Caption = "Current Project : " & gstr_Job & " , View Saved Data"
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
End If

sql = "Select member_sort from code_jis " & _
      "where member_type = 'Hbeam' " & _
      "group by member_sort " & _
      "order by member_sort "

Call Query_AddList2_function(0, cmbColumn, sql)
Call cmbColumn_Click

If gstr_MCEP_Flag = "Module01" Then
    WhereSql = "where (type = 'A1' " & _
               "or type = 'A2' " & _
               "or type = 'A3' " & _
               "or type = 'A4') "
ElseIf gstr_MCEP_Flag = "Module02" Then
    WhereSql = "where (type = 'B1' " & _
               "or type = 'B2' " & _
               "or type = 'B3' " & _
               "or type = 'B4' " & _
               "or type = 'B5') "
End If

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
Else
               sql = "Select Member_Name from MC_Connection " & WhereSql
               sql = sql & "and job = '" & gstr_Job & "'"
'               sql = sql & "in (select member_name from MC_Connection where job = '" & gstr_Job & "')"
               Call Query_AddList_function(1, lstLeftBeam, sql)
               Call Query_AddList_function(1, lstRightBeam, sql)
End If

Call gs_CobAddItem(cmbCode)

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
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub lstColumn_Click()

lblColumn.Caption = Trim(lstColumn.List(lstColumn.ListIndex))

End Sub
Private Sub lstBeam_Click()
lblBeam.Caption = Trim(lstBeam.List(lstBeam.ListIndex))

End Sub
Private Sub lstBracing_Click()
lblBracing.Caption = Trim(lstBracing.List(lstBracing.ListIndex))

End Sub

Private Sub Form_Control()

Select Case gstr_MCEP_type
    Case "Type01"
        fraLeftBeam.Enabled = False
        lstLeftBeam.Visible = False
        fraRightBeam.Enabled = True
        lstRightBeam.Visible = True
    Case "Type02", "Type05"
        fraLeftBeam.Enabled = True
        lstLeftBeam.Visible = True
        fraRightBeam.Enabled = False
        lstRightBeam.Visible = False
    Case Else
End Select

Call Type_BMP_Control_Start

End Sub
 
Private Sub Type_BMP_Control_Start()

    If gstr_MCdir_Flag = "Y" Then
        imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\YZ_Axis.bmp")
    Else
        imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\XZ_Axis.bmp")
    End If
    
    fraType.Caption = "View of " & gstr_MCEP_type & " (" & gstr_MCEP_Flag & ") "

    Select Case gstr_MCEP_Flag
        Case "Module01"
            Select Case gstr_MCEP_type
                Case "Type01"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_01.bmp")
                Case "Type02"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_02.bmp")
                Case "Type03"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_03.bmp")
                Case "Type04"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_04.bmp")
                Case "Type05"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_05.bmp")
            End Select
        Case "Module02"
            Select Case gstr_MCEP_type
                Case "Type01"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_06.bmp")
                Case "Type02"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_07.bmp")
                Case "Type03"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_08.bmp")
                Case "Type04"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_09.bmp")
                Case "Type05"
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\MO\Type\Type_10.bmp")
            End Select
    End Select

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

Private Sub lstLeftBeam_Click()

Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset
lblLeftBeam.Caption = Trim(lstLeftBeam.List(lstLeftBeam.ListIndex))

lstr_MemberName = CStr(Trim(lstLeftBeam.List(lstLeftBeam.ListIndex)))

xSQL = "Select Code from MC_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

lstr_Code_Left = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing


End Sub

Private Sub lstRightBeam_Click()

Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset
lblRightBeam.Caption = Trim(lstRightBeam.List(lstRightBeam.ListIndex))

lstr_MemberName = CStr(Trim(lstRightBeam.List(lstRightBeam.ListIndex)))

xSQL = "Select Code from MC_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

lstr_Code_Right = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing


End Sub

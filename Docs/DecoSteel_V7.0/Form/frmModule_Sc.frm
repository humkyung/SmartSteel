VERSION 5.00
Begin VB.Form frmMoudle_Sc 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmModule_Sc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10050
   Begin VB.Frame Frame2 
      Caption         =   "Selcet Code"
      Height          =   645
      Left            =   6720
      TabIndex        =   29
      Top             =   30
      Width           =   3285
      Begin VB.ComboBox cmbCode 
         Height          =   300
         Left            =   1290
         TabIndex        =   30
         Text            =   "JIS"
         Top             =   210
         Width           =   1935
      End
   End
   Begin VB.Frame fraColumn 
      Caption         =   "Select Column"
      Height          =   2310
      Left            =   30
      TabIndex        =   23
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstColumn 
         Columns         =   2
         Height          =   1500
         Left            =   90
         TabIndex        =   25
         Top             =   630
         Width           =   3165
      End
      Begin VB.ComboBox cmbColumn 
         Height          =   300
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.CommandButton cmbMake 
      Caption         =   "Make PML"
      Height          =   480
      Left            =   6840
      TabIndex        =   22
      Top             =   6630
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select PML Unit"
      Height          =   930
      Left            =   60
      TabIndex        =   16
      Top             =   6150
      Width           =   3255
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   600
         TabIndex        =   20
         Top             =   330
         Width           =   735
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   1635
         TabIndex        =   19
         Top             =   330
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   690
      End
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   1635
         TabIndex        =   17
         Top             =   600
         Width           =   600
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Select Modeling Option"
      Height          =   1320
      Left            =   30
      TabIndex        =   13
      Top             =   4770
      Width           =   3300
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   270
         TabIndex        =   15
         Text            =   "Type-01"
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Select Plate Type :"
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "View of Plate Type"
      Height          =   2340
      Left            =   3375
      TabIndex        =   12
      Top             =   4770
      Width           =   3300
      Begin VB.Image imgModel 
         Height          =   1950
         Left            =   315
         Picture         =   "frmModule_Sc.frx":000C
         Top             =   270
         Width           =   2745
      End
   End
   Begin VB.Frame fraModule 
      Caption         =   "View of Module-01"
      Height          =   2310
      Left            =   3375
      TabIndex        =   11
      Top             =   45
      Width           =   3300
      Begin VB.Image imgModule 
         Height          =   2025
         Left            =   315
         Picture         =   "frmModule_Sc.frx":1189E
         Top             =   225
         Width           =   2850
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "View Axis"
      Height          =   2310
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   3300
      Begin VB.Image imgAxis 
         Height          =   1905
         Left            =   225
         Picture         =   "frmModule_Sc.frx":24684
         Top             =   270
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Your Selection"
      Height          =   1590
      Left            =   6705
      TabIndex        =   5
      Top             =   765
      Width           =   3300
      Begin VB.Label Label3 
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
         TabIndex        =   27
         Top             =   240
         Width           =   1485
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
         Left            =   285
         TabIndex        =   26
         Top             =   420
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
         Left            =   285
         TabIndex        =   9
         Top             =   1275
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
         Left            =   285
         TabIndex        =   8
         Top             =   840
         Width           =   345
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
         TabIndex        =   7
         Top             =   1095
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
         TabIndex        =   6
         Top             =   660
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   8520
      TabIndex        =   4
      Top             =   6630
      Width           =   1455
   End
   Begin VB.Frame fraSubBeam 
      Caption         =   "Select Sub Beam"
      Height          =   2310
      Left            =   6705
      TabIndex        =   1
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstSubBeam 
         Columns         =   2
         Height          =   1860
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   3165
      End
   End
   Begin VB.Frame fraBeam 
      Caption         =   "Select Beam"
      Height          =   2310
      Left            =   3375
      TabIndex        =   0
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstBeam 
         Columns         =   2
         Height          =   1500
         Left            =   90
         TabIndex        =   2
         Top             =   630
         Width           =   3075
      End
      Begin VB.ListBox lstM01beam 
         Height          =   1860
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   3135
      End
      Begin VB.ComboBox cmbBeam 
         Height          =   300
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmMoudle_Sc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstr_Code As String
Dim gstr_Job As String, gs_PMLunit As String

Private Sub cmbCode_Click()
Dim lstr_Code As String
Dim xSQL As String

lstr_Code = CStr(Trim(cmbCode.Text))
    

If lstr_Code = "JIS" Then
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
               "where member_type = 'hbeam' " & _
               "order by member_no "
               cmbColumn.Enabled = True
               cmbBeam.Enabled = True
Else
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
              "where member_type = 'hbeam' "
               cmbColumn.Enabled = False
               cmbBeam.Enabled = False
               
End If
Call Query_AddList2_function(0, lstColumn, xSQL)
Call Query_AddList2_function(0, lstBeam, xSQL)
'Call Query_AddList2_function(0, lstM01beam, xSQL)

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

'Private Sub cmbSubBeam_Click()
'Dim sql As String
'
'If cmbSubBeam.Text <> "" Then
'    sql = "Select * from code_jis " & _
'          "where member_type = 'Hbeam' " & _
'          "and member_sort = '" & cmbSubBeam.Text & "' " & _
'          "order by member_no "
'Else
'    sql = "Select * from code_jis " & _
'          "where member_type = 'Hbeam' " & _
'          "order by member_no "
'End If
'
'Call Query_AddList_function(lstSubBeam, sql)
'
'End Sub

Private Sub cmbMake_Click()
Dim SubBeamName As String
Dim BeamName As String
Dim ColumnName As String
Dim tempColShape As String, tempType As String, TempUnit As String, TempPath As String
Dim formCode As String

On Error GoTo Labelstop

    ColumnName = lblColumn.Caption
    BeamName = lblBeam.Caption
    SubBeamName = lblSubBeam.Caption
    formCode = CStr(Trim(cmbCode.Text))
    If formCode = "" Then MsgBox "Code Selection Error. You must select Code Selection !!!": Exit Sub

    Select Case gstr_SCP_Flag
        Case "Module01"
            If ColumnName = "N/A" Then
                MsgBox "Column을 선택 하십시요."
                Exit Sub
            End If
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
        Case "Module02"
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
            If SubBeamName = "N/A" Then
                MsgBox "Sub Beam을 선택 하십시요."
                Exit Sub
            End If
    End Select
    
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

    If BeamName = "" Then
        MsgBox "Select Column Member Size....."
    Else
'        frmMain.CommonDialog.CancelError = True
        frmMain.CommonDialog.InitDir = App.Path
        frmMain.CommonDialog.DialogTitle = "Save PML File "
        frmMain.CommonDialog.Filter = "BasePlate (*.pml)|*.pml|"
        frmMain.CommonDialog.FileName = "Test.pml"
        
        frmMain.CommonDialog.ShowSave
        TempPath = frmMain.CommonDialog.FileName
    
        Call SC_PML(TempPath, gstr_Job, lstr_Code, formCode, gstr_SCdir_Flag, gstr_SCP_Flag, gstr_SCP_type, tempType, TempUnit, _
                    ColumnName, BeamName, SubBeamName)
        Call PML_Run(TempPath)
        End
    End If
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

'Me.Top = 0: Me.Left = 0
'lstBType.AddItem "Angle"
'lstBType.AddItem "Channel"
'lstBType.AddItem "Double Angle"
'lstBType.AddItem "Double Channel"
'lstBType.AddItem "Tee"
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
      fraModule.Caption = "Current Project : " & gstr_Job & " , View Saved Data"
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
End If


sql = "Select member_sort from code_jis " & _
      "where member_type = 'Hbeam' " & _
      "group by member_sort " & _
      "order by member_sort "

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

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
Else
               If gstr_SCP_Flag = "Module01" Then
                              Call Query_AddList2_function(0, cmbColumn, sql)
                              Call cmbColumn_Click
               '               Call Query_AddList2_function(cmbSubBeam, sql)
               '               Call cmbSubBeam_Click
                              sql = "Select Member_Name from SC_Connection where job = '" & gstr_Job & "'"
                              Call Query_AddList2_function(1, lstM01beam, sql)
               '               Call cmbBeam_Click
               Else
                              
                              Call Query_AddList2_function(0, cmbBeam, sql)
                              Call cmbBeam_Click
                              sql = "Select Member_Name from SC_Connection where job = '" & gstr_Job & "'"
                              Call Query_AddList2_function(1, lstSubBeam, sql)
               '               Call cmbSubBeam_Click
               
               End If
End If

Call gs_CobAddItem(cmbCode)


End Sub
'Private Sub lstBType_Click()
'Dim strTemp As String
'Dim sql As String
'
'strTemp = Trim(lstBType.List(lstBType.ListIndex))
'sql = "Select Member_Name from HB_Connection" & _
'      " where Shape = '" & strTemp & "'"
'
'Call Query_AddList_function(lstBracing, sql)

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

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

'End Sub

Private Sub lstColumn_Click()
lblColumn.Caption = Trim(lstColumn.List(lstColumn.ListIndex))

End Sub

Private Sub lstM01beam_Click()

Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset

lblBeam.Caption = Trim(lstM01beam.List(lstM01beam.ListIndex))

lstr_MemberName = CStr(Trim(lstM01beam.List(lstM01beam.ListIndex)))

xSQL = "Select Code from SC_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

lstr_Code = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing


End Sub

Private Sub lstSubBeam_Click()

Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset

lblSubBeam.Caption = Trim(lstSubBeam.List(lstSubBeam.ListIndex))

lstr_MemberName = CStr(Trim(lstSubBeam.List(lstSubBeam.ListIndex)))

xSQL = "Select Code from SC_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

lstr_Code = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing

End Sub
Private Sub lstBeam_Click()
lblBeam.Caption = Trim(lstBeam.List(lstBeam.ListIndex))

End Sub
Private Sub lstBracing_Click()
lblBracing.Caption = Trim(lstBracing.List(lstBracing.ListIndex))

End Sub

Private Sub Form_Control()

Select Case gstr_SCP_Flag
    Case "Module01"
        fraColumn.Enabled = True
        lstColumn.Visible = True
        cmbColumn.Visible = True
        fraBeam.Enabled = True
        lstM01beam.Visible = True
        lstBeam.Visible = False
        cmbBeam.Visible = False
        fraSubBeam.Enabled = False
        lstSubBeam.Visible = False
'        cmbSubBeam.Visible = False
        cmbType.Clear
        cmbType.AddItem "Type A"
        cmbType.AddItem "Type B"
        cmbType.ListIndex = 0
    Case "Module02"
        fraColumn.Enabled = False
        lstColumn.Visible = False
        cmbColumn.Visible = False
        lstM01beam.Visible = False
        fraBeam.Enabled = True
        lstBeam.Visible = True
        cmbBeam.Visible = True
        fraSubBeam.Enabled = True
        lstSubBeam.Visible = True
'        cmbSubBeam.Visible = True
        cmbType.Clear
        cmbType.AddItem "Type A"
        cmbType.AddItem "Type B"
        cmbType.AddItem "Type C"
        cmbType.AddItem "Type D"
        cmbType.ListIndex = 0
End Select


Call Type_BMP_Control_Start


End Sub

Private Sub Type_BMP_Control_Start()

    If gstr_SCdir_Flag = "Y" Then
        imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\YZ_Axis.bmp")
    Else
        imgAxis.Picture = LoadPicture(App.Path & "\BMP\VB\XZ_Axis.bmp")
    End If
    
    fraModule.Caption = "View of " & gstr_SCP_type & " (" & gstr_SCP_Flag & ") "

    Select Case gstr_SCP_Flag
        Case "Module01"
            Select Case gstr_SCP_type
                Case "Type01"
                    imgModule.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type01.bmp")
                Case "Type02"
                    imgModule.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type02.bmp")
            End Select
        Case "Module02"
            Select Case gstr_SCP_type
                Case "Type01"
                    imgModule.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type01.bmp")
                Case "Type02"
                    imgModule.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type02.bmp")
            End Select
    End Select

End Sub
Private Sub Type_BMP_Control()
Select Case gstr_SCP_Flag
    Case "Module01"
        Select Case gstr_SCP_type
            Case "Type01"
                If CStr(Trim(cmbType.Text)) = "Type A" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type01_A.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type B" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type01_B.bmp")
                End If
            Case "Type02"
                If CStr(Trim(cmbType.Text)) = "Type A" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type02_A.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type B" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul01_Type02_B.bmp")
                End If
        End Select
    Case "Module02"
        Select Case gstr_SCP_type
            Case "Type01"
                If CStr(Trim(cmbType.Text)) = "Type A" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type01_A.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type B" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type01_B.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type C" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type01_C.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type D" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type01_D.bmp")
                End If
            Case "Type02"
                If CStr(Trim(cmbType.Text)) = "Type A" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type02_A.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type B" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type02_B.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type C" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type02_C.bmp")
                ElseIf CStr(Trim(cmbType.Text)) = "Type D" Then
                    imgModel.Picture = LoadPicture(App.Path & "\BMP\SC\SC_Modul02_Type02_D.bmp")
                End If
        End Select
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


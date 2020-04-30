VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   1  '단일 고정
   Caption         =   "General"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame4 
      Caption         =   "Job Select"
      Height          =   3795
      Left            =   6090
      TabIndex        =   25
      Top             =   0
      Width           =   2385
      Begin VB.ListBox lstJob 
         Height          =   2940
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   2145
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Select"
         Height          =   405
         Left            =   180
         TabIndex        =   26
         Top             =   3300
         Width           =   2085
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete All"
      Height          =   465
      Left            =   4560
      TabIndex        =   24
      Top             =   2790
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   585
      Left            =   30
      TabIndex        =   20
      Top             =   -30
      Width           =   4455
      Begin VB.CommandButton cmdJobSlect 
         Caption         =   "Job Select"
         Height          =   315
         Left            =   3300
         TabIndex        =   23
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox txtJobName 
         Height          =   285
         Left            =   1380
         TabIndex        =   22
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Job Name :"
         Height          =   180
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   4500
      TabIndex        =   10
      Top             =   -30
      Width           =   1545
      Begin VB.CheckBox chk02 
         Caption         =   "Window On Bottom"
         Enabled         =   0   'False
         Height          =   495
         Left            =   180
         TabIndex        =   12
         Top             =   870
         Width           =   1185
      End
      Begin VB.CheckBox chk01 
         Caption         =   "Window OnTop"
         Height          =   405
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   465
      Left            =   4560
      TabIndex        =   1
      Top             =   2250
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   4560
      TabIndex        =   0
      Top             =   3330
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   4455
      Begin VB.Frame Frame5 
         Caption         =   "Select PML Unit"
         Height          =   630
         Left            =   180
         TabIndex        =   28
         Top             =   2520
         Width           =   4155
         Begin VB.OptionButton optFeet 
            Caption         =   "feet"
            Height          =   240
            Left            =   3015
            TabIndex        =   32
            Top             =   300
            Width           =   600
         End
         Begin VB.OptionButton optInch 
            Caption         =   "inch"
            Height          =   180
            Left            =   2040
            TabIndex        =   31
            Top             =   300
            Width           =   690
         End
         Begin VB.OptionButton optM 
            Caption         =   "m"
            Height          =   195
            Left            =   1245
            TabIndex        =   30
            Top             =   300
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optMM 
            Caption         =   "mm"
            Height          =   240
            Left            =   330
            TabIndex        =   29
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.TextBox txtGrade 
         Height          =   270
         Left            =   2850
         TabIndex        =   19
         Text            =   "A36"
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txtMaterial 
         Height          =   270
         Left            =   2850
         TabIndex        =   18
         Text            =   "Steel"
         Top             =   525
         Width           =   1545
      End
      Begin VB.ComboBox cmbBPClass 
         Height          =   300
         ItemData        =   "frmGeneral.frx":0000
         Left            =   2850
         List            =   "frmGeneral.frx":0002
         Style           =   2  '드롭다운 목록
         TabIndex        =   17
         Top             =   840
         Width           =   1545
      End
      Begin VB.ComboBox cmbHBClass 
         Height          =   300
         ItemData        =   "frmGeneral.frx":0004
         Left            =   2850
         List            =   "frmGeneral.frx":0006
         Style           =   2  '드롭다운 목록
         TabIndex        =   16
         Top             =   1500
         Width           =   1545
      End
      Begin VB.ComboBox cmbVBClass 
         Height          =   300
         ItemData        =   "frmGeneral.frx":0008
         Left            =   2850
         List            =   "frmGeneral.frx":000A
         Style           =   2  '드롭다운 목록
         TabIndex        =   15
         Top             =   1155
         Width           =   1545
      End
      Begin VB.ComboBox cmbMCClass 
         Height          =   300
         ItemData        =   "frmGeneral.frx":000C
         Left            =   2850
         List            =   "frmGeneral.frx":000E
         Style           =   2  '드롭다운 목록
         TabIndex        =   14
         Top             =   1830
         Width           =   1545
      End
      Begin VB.ComboBox cmbSCClass 
         Height          =   300
         ItemData        =   "frmGeneral.frx":0010
         Left            =   2850
         List            =   "frmGeneral.frx":0012
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   2145
         Width           =   1545
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Shear Connection Class"
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   2190
         Width           =   2070
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Moment Connection Class"
         Height          =   180
         Left            =   300
         TabIndex        =   8
         Top             =   1890
         Width           =   2265
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Horizontal Bracing Class"
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   1560
         Width           =   2100
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vertical Bracing Class"
         Height          =   180
         Left            =   300
         TabIndex        =   6
         Top             =   1230
         Width           =   1890
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Base Plate Class"
         Height          =   180
         Left            =   300
         TabIndex        =   5
         Top             =   900
         Width           =   1470
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Material"
         Height          =   180
         Left            =   300
         TabIndex        =   4
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grade"
         Height          =   180
         Left            =   300
         TabIndex        =   3
         Top             =   270
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chk01_Click()
Dim i
If chk01.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If chk02.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub chk01_gotFocus()
chk02.Value = 0
gin_Chk_Flag01 = 1
gin_Chk_Flag02 = 0

End Sub

Private Sub chk02_Click()
Dim i
If chk01.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
If chk02.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub

Private Sub chk02_gotFocus()
chk01.Value = 0
gin_Chk_Flag02 = 1
gin_Chk_Flag01 = 0

End Sub

Private Sub cmdDelete_Click()
Dim xSQL As String, Response As String

Response = MsgBox("All Job Information will be Deleted. Are you sure ?", vbYesNo)

If Response = vbYes Then
               xSQL = "delete from Plate_General where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from BasePlate_Fixed where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from BasePlate_Hinged where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from HB_Connection where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from VB_Connection where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from MC_Connection where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               xSQL = "delete from SC_Connection where job = '" & Trim(txtJobName.Text) & "'"
               Set reData1 = adoConnection1.Execute(xSQL)
               Set reData1 = Nothing
               
               txtJobName.Text = ""
               txtGrade.Text = ""
               txtMaterial.Text = ""
               cmbBPClass.Text = "1"
               cmbVBClass.Text = "1"
               cmbHBClass.Text = "1"
               cmbMCClass.Text = "1"
               cmbSCClass.Text = "1"
               
               frmMain.StatusBar1.Panels(1).Text = "Current Job is Deleted....."
End If
End Sub

Private Sub cmdExit_Click()
Dim i As Integer
For i = CInt(Trim(Me.Height)) To 0 Step -40
               Me.Width = i
               Me.Height = i
Next i

If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If

End Sub

Private Sub cmdJobSlect_Click()
'frmJobSelect.Show
Dim i As Integer
'Me.Height = 3540
'
'Me.Width = 6180

'For i = CInt(Trim(Me.Width)) To 8610 Step 1
Frame4.Top = -5
Frame4.Left = 4500

For i = CInt(Trim(Me.Width)) To 7000 Step 1
               Me.Width = i
Next i



'Me.Top = 0
'Me.Left = 0
Dim xSQL As String
xSQL = "Select job from Plate_General"
Call Query_AddList2_function(1, lstJob, xSQL)

End Sub

Private Sub cmdOK_Click()
'Me.Top = 0: Me.Left = 0
'Me.Width = 6180
'Me.Height = 3540

Frame4.Left = 6090

On Error GoTo Erro100

xSQL = "select job,Grade,Material,BP_Class, VB_Class, HB_Class, MC_Class, SC_Class from Plate_General where job = '"
xSQL = xSQL & CStr(Trim(lstJob.List(lstJob.ListIndex))) & "'"
Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    txtJobName.Text = reData1!job
    txtGrade.Text = reData1!grade
    txtMaterial.Text = reData1!material
    cmbBPClass.Text = reData1!bp_class
    cmbVBClass.Text = reData1!vb_class
    cmbHBClass.Text = reData1!hb_class
    cmbMCClass.Text = reData1!mc_class
    cmbSCClass.Text = reData1!sc_class
End If
reData1.Close
Set reData1 = Nothing

'gstr_Job = CStr(Trim(txtJobName.Text))

If CStr(Trim(txtJobName.Text)) = "" Then
               frmMain.StatusBar1.Panels(1).Text = "Current Job is not selected....."
Else
               frmMain.StatusBar1.Panels(1).Text = "Current Job : " & gstr_Job
End If


For i = CInt(Trim(Me.Width)) To 6180 Step -1
      Me.Width = i
Next i

'If Len(Command) <> 0 Then
Call gs_project_Output(CStr(Trim(txtJobName.Text)))
'End If

Dim TempUnit As String
Open App.Path & "\" & Trim(txtJobName.Text) & "_pmlunit.ini" For Input As #1
      Input #1, TempUnit
Close #1


If TempUnit = "mm" Then
      optMM.Value = True
      optM.Value = False
      optInch.Value = False
      optFeet.Value = False
ElseIf TempUnit = "m" Then
      optMM.Value = False
      optM.Value = True
      optInch.Value = False
      optFeet.Value = False
ElseIf TempUnit = "inch" Then
      optMM.Value = False
      optM.Value = False
      optInch.Value = True
      optFeet.Value = False
ElseIf TempUnit = "feet" Then
      optMM.Value = False
      optM.Value = False
      optInch.Value = False
      optFeet.Value = True
End If

'gs_PMLunit = TempUnit

Exit Sub
Erro100:
MsgBox "Make (project_name)_pmlunit.ini File at c:\decosteel"

End Sub

Private Sub cmdSave_Click()
Dim xSQL As String
Dim TempUnit As String

If CStr(Trim(txtJobName.Text)) = "" Then
            MsgBox "Project Name Miss........."
            Exit Sub
End If

Screen.MousePointer = vbHourglass

If cmbBPClass.Text = "" Then
      MsgBox "Base Plate Class Miss........."
      Screen.MousePointer = vbDefault
      Exit Sub
End If

If cmbVBClass.Text = "" Then
      MsgBox "Vertical Bracing Class Miss........."
      Screen.MousePointer = vbDefault
      Exit Sub
End If

If cmbHBClass.Text = "" Then
      MsgBox "Horizontal Bracing Class Miss........."
      Screen.MousePointer = vbDefault
      Exit Sub
End If

If cmbMCClass.Text = "" Then
      MsgBox "Moment Connection Class Miss........."
      Screen.MousePointer = vbDefault
      Exit Sub
End If

If cmbSCClass.Text = "" Then
      MsgBox "Shear Connection Class Miss........."
      Screen.MousePointer = vbDefault
      Exit Sub
End If

If optMM.Value = True Then
      TempUnit = "mm"
ElseIf optM.Value = True Then
      TempUnit = "m"
ElseIf optInch.Value = True Then
      TempUnit = "inch"
ElseIf optFeet.Value = True Then
      TempUnit = "feet"
End If


xSQL = "delete from Plate_General where job = '" & Trim(txtJobName.Text) & "'"
Set reData1 = adoConnection1.Execute(xSQL)

xSQL = "insert into Plate_General ("
xSQL = xSQL & "job, "
xSQL = xSQL & "grade, "
xSQL = xSQL & "material, "
xSQL = xSQL & "bp_class, "
xSQL = xSQL & "vb_class, "
xSQL = xSQL & "hb_class, "
xSQL = xSQL & "mc_class, "
xSQL = xSQL & "sc_class) values ("
xSQL = xSQL & "'" & Trim(txtJobName.Text) & "', "
xSQL = xSQL & "'" & Trim(txtGrade.Text) & "', "
xSQL = xSQL & "'" & Trim(txtMaterial.Text) & "', "
xSQL = xSQL & "'" & Trim(cmbBPClass.Text) & "', "
xSQL = xSQL & "'" & Trim(cmbVBClass.Text) & "', "
xSQL = xSQL & "'" & Trim(cmbHBClass.Text) & "', "
xSQL = xSQL & "'" & Trim(cmbMCClass.Text) & "', "
xSQL = xSQL & "'" & Trim(cmbSCClass.Text) & "' "
xSQL = xSQL & ")"

adoConnection1.Execute (xSQL)

Set reData1 = Nothing
Screen.MousePointer = vbDefault

'If Len(Command) <> 0 Then
            Call gs_project_Output(CStr(Trim(txtJobName.Text)))
'End If

Open App.Path & "\" & Trim(txtJobName.Text) & "_pmlunit.ini" For Output As #1
      Print #1, TempUnit
Close #1


End Sub

Private Sub Form_Load()
Dim xSQL As String

Me.Top = 0: Me.Left = 0

Me.Width = 6180
Me.Height = 4260

'chk01.Value = 1

If Len(Command) <> 0 Then
      chk01.Enabled = False
      chk02.Enabled = False
      If gin_Chk_Flag01 = 1 Then chk01.Value = 1
      If gin_Chk_Flag02 = 1 Then chk02.Value = 1
Else
      chk01.Value = 1
      chk02.Enabled = True
End If

If gin_Chk_Flag01 = 0 Then
            i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
            i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

txtJobName.Text = gstr_Job
xSQL = "select Grade,Material,BP_Class, VB_Class, HB_Class, MC_Class, SC_Class from Plate_General where job = '"
xSQL = xSQL & CStr(Trim(txtJobName.Text)) & "'"
Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    txtGrade.Text = reData1!grade
    txtMaterial.Text = reData1!material
    cmbBPClass.Text = reData1!bp_class
    cmbVBClass.Text = reData1!vb_class
    cmbHBClass.Text = reData1!hb_class
    cmbMCClass.Text = reData1!mc_class
    cmbSCClass.Text = reData1!sc_class
End If
reData1.Close
Set reData1 = Nothing

cmbBPClass.AddItem "0"
cmbBPClass.AddItem "1"
cmbBPClass.AddItem "2"
cmbBPClass.AddItem "3"
cmbBPClass.AddItem "4"

cmbVBClass.AddItem "0"
cmbVBClass.AddItem "1"
cmbVBClass.AddItem "2"
cmbVBClass.AddItem "3"
cmbVBClass.AddItem "4"

cmbHBClass.AddItem "0"
cmbHBClass.AddItem "1"
cmbHBClass.AddItem "2"
cmbHBClass.AddItem "3"
cmbHBClass.AddItem "4"

cmbMCClass.AddItem "0"
cmbMCClass.AddItem "1"
cmbMCClass.AddItem "2"
cmbMCClass.AddItem "3"
cmbMCClass.AddItem "4"

cmbSCClass.AddItem "0"
cmbSCClass.AddItem "1"
cmbSCClass.AddItem "2"
cmbSCClass.AddItem "3"
cmbSCClass.AddItem "4"


End Sub


Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If 'Dim i
'If chk01.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
'If chk02.Value = 1 Then i = SetWindowPos(frmMain.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
'gstr_Job = CStr(Trim(txtJobName.Text))

End Sub


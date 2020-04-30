VERSION 5.00
Begin VB.Form frmJobSelect 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Job Select"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "frmJobSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   2640
      Width           =   2265
   End
   Begin VB.ListBox lstJob 
      Height          =   2400
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   2265
   End
End
Attribute VB_Name = "frmJobSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

xSQL = "select job,Grade,Material,BP_Class, VB_Class, HB_Class, MC_Class, SC_Class from Plate_General where job = '"
xSQL = xSQL & CStr(Trim(lstJob.List(lstJob.ListIndex))) & "'"
Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    frmGeneral.txtJobName.Text = reData1!job
    frmGeneral.txtGrade.Text = reData1!grade
    frmGeneral.txtMaterial.Text = reData1!material
    frmGeneral.cmbBPClass.Text = reData1!bp_class
    frmGeneral.cmbVBClass.Text = reData1!vb_class
    frmGeneral.cmbHBClass.Text = reData1!hb_class
    frmGeneral.cmbMCClass.Text = reData1!mc_class
    frmGeneral.cmbSCClass.Text = reData1!sc_class
End If
reData1.Close
Set reData1 = Nothing

Unload Me
End Sub

Private Sub Form_Load()
Dim xSQL As String
xSQL = "Select job from Plate_General"

Call Query_AddList2_function(1, lstJob, xSQL)



End Sub


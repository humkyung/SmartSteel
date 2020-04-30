VERSION 5.00
Begin VB.Form frmTableDrop 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Code Table Drop"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   2295
      TabIndex        =   2
      Top             =   540
      Width           =   1590
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "Table Drop"
      Height          =   420
      Left            =   2295
      TabIndex        =   1
      Top             =   45
      Width           =   1590
   End
   Begin VB.ListBox lstDrop 
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2220
   End
End
Attribute VB_Name = "frmTableDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDrop_Click()
Dim sql As String
Dim DropTable As String, str As String
Dim strArray() As String, intCount As Integer
Dim rstSchema As ADODB.Recordset

On Error GoTo Labelstop

DropTable = CStr(Trim(lstDrop.List(lstDrop.ListIndex)))
sql = "DROP TABLE " & DropTable
adoConnection.Execute (sql)

strArray = Split(DropTable, "_")

sql = "Delete from code where codename = '" & strArray(1) & "'"
adoConnection.Execute (sql)
MsgBox "Table Drop Success......"

lstDrop.Clear
Set rstSchema = adoConnection.OpenSchema(adSchemaTables)
Do Until rstSchema.EOF
            str = CStr(Trim(rstSchema!TABLE_NAME))
            strArray = Split(str, "_")
            intCount = UBound(strArray, 1)
            If intCount = 1 Then
                        If strArray(1) <> "JIS" And strArray(1) <> "AISC" Then
                                    lstDrop.AddItem str
                        End If
            End If
            rstSchema.MoveNext
Loop
rstSchema.Close
Exit Sub


Labelstop:
MsgBox "Table Drop Fail........"
End Sub

Private Sub cmdExit_Click()
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If
End Sub

Private Sub Form_Load()
Dim str As String
Dim rstSchema As ADODB.Recordset
Dim intCount As Integer
Dim strArray() As String

If gin_Chk_Flag01 = 0 Then
            i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
            i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

'If Len(Command) <> 0 Then
'            Dim lstr_ProjectName As String, lin_Error As Integer
'            Call gs_Call_Project(lstr_ProjectName, lin_Error)
'            If lin_Error = 1 Then MsgBox "Job is not selected. Please Job Select at General": End
'            gstr_Job = lstr_ProjectName
'            Me.Caption = "Current Project : " & gstr_Job & " , Code Drop"
'End If

Set rstSchema = adoConnection.OpenSchema(adSchemaTables)

Do Until rstSchema.EOF
            str = CStr(Trim(rstSchema!TABLE_NAME))
            strArray = Split(str, "_")
            intCount = UBound(strArray, 1)
            If intCount = 1 Then
                        If strArray(1) <> "JIS" And strArray(1) <> "AISC" Then
                                    lstDrop.AddItem str
                        End If
            End If
            rstSchema.MoveNext
Loop
rstSchema.Close
   
  
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If
End Sub

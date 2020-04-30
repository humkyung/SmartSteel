VERSION 5.00
Begin VB.Form frmMember 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Select Name of Steel Member"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
   Icon            =   "frmMember.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ComboBox cmbCode 
      Height          =   300
      Left            =   1620
      TabIndex        =   3
      Text            =   "JIS"
      Top             =   60
      Width           =   1155
   End
   Begin VB.ComboBox cmbQuery 
      Height          =   300
      Left            =   1620
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   360
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O.K"
      Height          =   405
      Left            =   690
      TabIndex        =   1
      Top             =   3900
      Width           =   1335
   End
   Begin VB.ListBox lstMember 
      Height          =   3120
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   2715
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Query Condition :"
      Height          =   180
      Left            =   30
      TabIndex        =   5
      Top             =   420
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code : "
      Height          =   180
      Left            =   30
      TabIndex        =   4
      Top             =   90
      Width           =   630
   End
End
Attribute VB_Name = "frmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstr_Code As String
Private Sub cmbCode_Click()
Dim xSQL As String
    lstr_Code = CStr(Trim(cmbCode.Text))
    

If lstr_Code = "JIS" Then
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
               "where member_type = '" & gstr_Shape & "' " & _
               "order by member_no "
               cmbQuery.Enabled = True
Else
               xSQL = "Select member_name from code_" & lstr_Code & " " & _
               "where member_type = '" & gstr_Shape & "' "
               cmbQuery.Enabled = False
End If
Call Query_AddList2_function(0, lstMember, xSQL)

    
End Sub

Private Sub cmbQuery_Click()
Dim sql As String
If cmbQuery.Text <> "" Then
    sql = "Select * from code_" & lstr_Code & " " & _
          "where member_type = 'hbeam' " & _
          "and member_sort = '" & cmbQuery.Text & "' " & _
          "order by member_no "
Else
    sql = "Select * from code_" & lstr_Code & " " & _
          "where member_type = 'hbeam' " & _
          "order by member_no "
End If

Call Query_AddList_function(0, lstMember, sql)


End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim xSQL As String

'cmbCode.AddItem "JIS"
'cmbCode.AddItem "AISC"
If gin_Chk_Flag01 = 0 Then
            i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
            i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

Call gs_CobAddItem(cmbCode)

lstr_Code = "JIS"
If gstr_Shape <> "Hbeam" Then
    cmbQuery.Enabled = False
Else

    xSQL = "Select member_sort from code_" & lstr_Code & " " & _
          "where member_type = 'hbeam' " & _
          "group by member_sort " & _
          "order by member_sort "
    
    Call Query_AddList2_function(0, cmbQuery, xSQL)
End If

xSQL = "Select member_name from code_" & lstr_Code & " " & _
      "where member_type = '" & gstr_Shape & "' " & _
      "order by member_no "

Call Query_AddList2_function(0, lstMember, xSQL)


End Sub

Private Sub lstMember_Click()
Select Case gin_Shape_Flag
    Case 1
        frmSpecVB.txtMName.Text = Trim(lstMember.List(lstMember.ListIndex))
        If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
        frmSpecVB.txtCode.Text = CStr(Trim(cmbCode.Text))
    Case 2
        frmSpecHB.txtMName.Text = Trim(lstMember.List(lstMember.ListIndex))
        If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
        frmSpecHB.txtCode.Text = CStr(Trim(cmbCode.Text))
    Case 3
        frmSpecHBP.txtMName.Text = Trim(lstMember.List(lstMember.ListIndex))
        If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
        frmSpecHBP.txtCode.Text = CStr(Trim(cmbCode.Text))
    Case 4
        frmSpecMC.txtMName.Text = Trim(lstMember.List(lstMember.ListIndex))
        If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
        frmSpecMC.txtCode.Text = CStr(Trim(cmbCode.Text))
    Case 5
       frmSpecSC.txtMName.Text = Trim(lstMember.List(lstMember.ListIndex))
       If CStr(Trim(cmbCode.Text)) = "" Then MsgBox "Select Code Name !!!": Exit Sub
        frmSpecSC.txtCode.Text = CStr(Trim(cmbCode.Text))
End Select

End Sub

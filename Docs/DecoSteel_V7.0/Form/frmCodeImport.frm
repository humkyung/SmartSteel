VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmCodeImport 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Code Import"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Make Table"
      Height          =   510
      Left            =   2520
      TabIndex        =   1
      Top             =   3780
      Width           =   2265
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "File Open"
      Height          =   510
      Left            =   120
      TabIndex        =   9
      Top             =   3780
      Width           =   2265
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   510
      Left            =   4950
      TabIndex        =   8
      Top             =   3780
      Width           =   2265
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   -45
      Width           =   3300
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   1350
         Style           =   2  '드롭다운 목록
         TabIndex        =   7
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code Unit :"
         Height          =   180
         Left            =   225
         TabIndex        =   6
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   -45
      Width           =   3930
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   1890
         TabIndex        =   4
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Input Code Name : "
         Height          =   180
         Left            =   225
         TabIndex        =   3
         Top             =   315
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7260
      Begin FPSpread.vaSpread SpreadSheet 
         Height          =   2775
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   7155
         _Version        =   393216
         _ExtentX        =   12621
         _ExtentY        =   4895
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
         MaxCols         =   7
         SpreadDesigner  =   "frmCodeImport.frx":0000
      End
   End
End
Attribute VB_Name = "frmCodeImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If
End Sub

Private Sub cmdOpen_Click()
On Error GoTo Labelstop
Dim lvPath As String

frmMain.CommonDialog.InitDir = App.Path
frmMain.CommonDialog.DialogTitle = "Open TXT File "
frmMain.CommonDialog.Filter = "*.txt"
'frmMain.CommonDialog.FileName = "Test.pml"

frmMain.CommonDialog.ShowOpen
lvPath = frmMain.CommonDialog.FileName

Call gsubSS_Clear(SpreadSheet)
Call ls_Data_Insert(lvPath)


Labelstop:
End Sub

Private Sub Command1_Click()
Dim sql As String
Dim i As Long
Dim convertL As Single, convertA As Single
Dim ls_Unit As String
Dim ls_Table As String
Dim ls_Type As String, ls_Name As String, ls_Depth As String, ls_Width As String, _
         ls_Tf As String, ls_Tw As String, ls_Area As String
Dim lsn_Depth As Single, lsn_Width As Single, lsn_Tf As Single, lsn_Tw As Single, lsn_Area As Single

On Error GoTo Labelstop

ls_Unit = CStr(Trim(cboUnit.Text))
ls_Table = "Code_" & CStr(Trim(txtCode.Text))

If CStr(Trim(txtCode.Text)) = "" Then
            MsgBox "Code Name Miss........"
            Exit Sub
End If
 
 If ls_Unit = "" Then
            MsgBox "Code Unit Miss........."
            Exit Sub
 End If
 
sql = "CREATE TABLE " & ls_Table & _
          " (Member_Type TEXT, Member_Name TEXT, Ax TEXT, D TEXT, Bf TEXT, Tf TEXT, Tw TEXT)"

adoConnection.Execute (sql)

sql = "Insert into code values('" & CStr(Trim(txtCode.Text)) & "')"
adoConnection.Execute (sql)


Select Case ls_Unit
            Case "inch"
                        convertL = 25.4
                        covertA = 6.4516
            Case "mm"
                        convertL = 1
                        convertA = 0.01
End Select

i = 1
Do While gfunSS_GetText(SpreadSheet, i, 1) <> ""
            ls_Type = gfunSS_GetText(SpreadSheet, i, 1)
            ls_Name = gfunSS_GetText(SpreadSheet, i, 2)
            ls_Depth = gfunSS_GetText(SpreadSheet, i, 3)
            ls_Width = gfunSS_GetText(SpreadSheet, i, 4)
            ls_Tf = gfunSS_GetText(SpreadSheet, i, 5)
            ls_Tw = gfunSS_GetText(SpreadSheet, i, 6)
            ls_Area = gfunSS_GetText(SpreadSheet, i, 7)
            
            If ls_Tw = "" Then ls_Tw = "0"
            
            lsn_Depth = CSng(ls_Depth) * convertL
            lsn_Width = CSng(ls_Width) * convertL
            lsn_Tf = CSng(ls_Tf) * convertL
            lsn_Tw = CSng(ls_Tw) * convertL
            lsn_Area = CSng(ls_Area) * convertA
            
            sql = "Insert into " & ls_Table & " values ( '"
            sql = sql & ls_Type & "', '"
            sql = sql & ls_Name & "', '"
            sql = sql & lsn_Area & "', '"
            sql = sql & lsn_Depth & "', '"
            sql = sql & lsn_Width & "', '"
            sql = sql & lsn_Tf & "', '"
            sql = sql & ls_Tw & "')"
            
            adoConnection.Execute (sql)
            
            i = i + 1
Loop
MsgBox "Table Creation Success.........."
Command1.Enabled = False
Exit Sub

Labelstop:
MsgBox "Table Creation fail........."
End Sub

Private Sub ls_Data_Insert(ByVal valPathName As String)
Dim ls_String01 As String, ls_String02 As String, ls_String03 As String, ls_String04 As String, ls_String05 As String, _
    ls_String06 As String, ls_String07 As String
Dim i As Long
Dim intCount As Integer, intCount02 As Integer
Dim strArray() As String, strArray02() As String, strArray03() As String
Dim tempWidth As Single

i = 1
Open valPathName For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount = 1 Then
                                    Input #1, ls_String02
                                    Input #1, ls_String03
                                    Input #1, ls_String04
                                    Input #1, ls_String05
                                    Input #1, ls_String06
                                    Input #1, ls_String07
                                    strArray02 = Split(ls_String02)
                                    strArray03 = Split(ls_String07)
                                    
                                    intCount = UBound(strArray02, 1)
                                    
                                    Select Case Trim(strArray(0))
                                                Case "1"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Hbeam")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, strArray02(1))
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, strArray02(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                Case "2"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Channel")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, strArray02(1))
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, strArray02(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                Case "3"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Tee")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, strArray02(1))
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, strArray02(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                Case "4"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Angle")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, strArray02(1))
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, "")
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                            
                                                Case "5"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Double Channel")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            
                                                            tempWidth = CSng(strArray02(1)) * 2 + CSng(strArray03(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, tempWidth)
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, strArray02(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                Case "6"
                                                            Call gsubSS_SetText(SpreadSheet, i, 1, "Double Angle")
                                                            Call gsubSS_SetText(SpreadSheet, i, 2, strArray(1))
                                                            
                                                            Call gsubSS_SetText(SpreadSheet, i, 3, strArray02(0))
                                                            
                                                            tempWidth = CSng(strArray02(1)) * 2 + CSng(strArray03(2))
                                                            Call gsubSS_SetText(SpreadSheet, i, 4, tempWidth)
                                                            Call gsubSS_SetText(SpreadSheet, i, 5, strArray02(4))
                                                            Call gsubSS_SetText(SpreadSheet, i, 6, "")
                                                            Call gsubSS_SetText(SpreadSheet, i, 7, strArray02(3))
                                                            i = i + 1
                                                Case Else
                                    End Select
                                    
                        End If
            Loop
Close #1

End Sub

Private Sub Form_Load()

If gin_Chk_Flag01 = 0 Then
      i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
      i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

'If Len(Command) <> 0 Then
'      Dim lstr_ProjectName As String, lin_Error As Integer
'      Call gs_Call_Project(lstr_ProjectName, lin_Error)
'      If lin_Error = 1 Then MsgBox "Job is not selected. Please Job Select at General": End
'      gstr_Job = lstr_ProjectName
'      Me.Caption = "Current Project : " & gstr_Job & " , Code Import"
'End If

Command1.Enabled = True
cboUnit.AddItem "mm"
cboUnit.AddItem "inch"

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Then
            Unload Me
Else
            End
End If
End Sub


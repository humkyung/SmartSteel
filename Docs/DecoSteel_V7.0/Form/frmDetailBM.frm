VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailBM 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Bolt & Plate Bill of Material"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmDetailBM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6555
   Begin MSComctlLib.ProgressBar PBar2 
      Height          =   375
      Left            =   990
      TabIndex        =   4
      Top             =   450
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   375
      Left            =   990
      TabIndex        =   3
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Frame Frame3 
      Caption         =   "Plate BOM "
      Height          =   4830
      Left            =   3285
      TabIndex        =   2
      Top             =   990
      Width           =   3210
      Begin MSComctlLib.ListView lstPlate_BM 
         Height          =   4560
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   8043
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Quantity (Unit : kg)"
            Object.Width           =   3440
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "High Tension Bolt BOM "
      Height          =   2355
      Left            =   45
      TabIndex        =   1
      Top             =   3465
      Width           =   3210
      Begin MSComctlLib.ListView lstHTB_BM 
         Height          =   2040
         Left            =   90
         TabIndex        =   9
         Top             =   225
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Quantity (Unit : EA)"
            Object.Width           =   3440
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Anchor Bolt BOM "
      Height          =   2400
      Left            =   45
      TabIndex        =   0
      Top             =   990
      Width           =   3210
      Begin MSComctlLib.ListView lstAB_BM 
         Height          =   2130
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3757
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Quantity (Unit : EA)"
            Object.Width           =   3440
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   960
      Left            =   0
      TabIndex        =   5
      Top             =   -90
      Width           =   6495
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open File"
         Height          =   780
         Left            =   45
         Picture         =   "frmDetailBM.frx":000C
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   135
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmDetailBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AB_BM() As String
Dim HTB_BM() As String
Dim Plate_BM() As String
Dim AB_Count As Integer
Dim HTB_Count As Integer
Dim Plate_Count As Integer

Private Sub cmdOpen_Click()
On Error GoTo Labelstop
Dim lvPath As String
Dim i As Integer
Dim L As ListItem

lstAB_BM.ListItems.Clear: lstHTB_BM.ListItems.Clear: lstPlate_BM.ListItems.Clear


frmMain.CommonDialog.InitDir = App.Path
frmMain.CommonDialog.DialogTitle = "Open TXT File "
frmMain.CommonDialog.Filter = "*.txt"
'frmMain.CommonDialog.FileName = "Test.pml"

frmMain.CommonDialog.ShowOpen
lvPath = frmMain.CommonDialog.FileName
PBar1.Visible = True
Frame4.Visible = False
Call ls_BM_Data_Insert(lvPath)

For i = 1 To AB_Count
            Set L = lstAB_BM.ListItems.Add(, , AB_BM(i, 1))
             L.SubItems(1) = AB_BM(i, 2)
Next i



For i = 1 To HTB_Count
            Set L = lstHTB_BM.ListItems.Add(, , HTB_BM(i, 1))
             L.SubItems(1) = HTB_BM(i, 2)
Next i


Dim Temp As String

For i = 1 To Plate_Count
            Temp = CStr(CSng(Plate_BM(i, 1)) * 1000) & " (t)"
            Set L = lstPlate_BM.ListItems.Add(, , Temp)
             L.SubItems(1) = Format(Plate_BM(i, 2), "0.00")
Next i



Labelstop:
PBar1.Visible = False
PBar2.Visible = False
Frame4.Visible = True
End Sub
Private Sub ls_BM_Data_Insert(ByVal valPath As String)
Dim i As Integer
For i = 1 To 500
            PBar1.Value = i
Next i

Call ls_AnchorBolt_BM(valPath)

For i = 501 To 1000
            PBar1.Value = i
Next i
Call ls_HTB_BM(valPath)

PBar2.Visible = True

For i = 1 To 500
            PBar2.Value = i
Next i

Call ls_Plate_BM(valPath)

For i = 501 To 1000
            PBar2.Value = i
Next i
End Sub

Private Sub ls_AnchorBolt_BM(ByVal valPath As String)

Dim ls_String01 As String, ls_String02 As String
Dim i As Integer
Dim intCount As Integer, intCount02 As Integer
Dim strArray() As String, strArray02() As String
Dim strAB_Size() As String, intAB As Integer
intAB = 0
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount = 35 Then
                                    strArray02 = Split(strArray(0), "_")
                                    If strArray02(0) = "AB" Then
                                                intAB = intAB + 1
                                    End If
                        End If
            Loop
Close #1
ReDim strAB_Size(1 To intAB) As String
i = 1
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount = 35 Then
                                    strArray02 = Split(strArray(0), "_")
                                    If strArray02(0) = "AB" Then
                                                strAB_Size(i) = strArray02(1)
                                                i = i + 1
                                    End If
                        End If
            Loop
Close #1

Dim sql As String
sql = "Delete from Anchorbolt_BM"
adoConnection1.Execute (sql)

For i = 1 To intAB
            sql = "Insert into AnchorBolt_BM values ('" & strAB_Size(i) & "')"
            adoConnection1.Execute (sql)
Next i

Dim local_adoRecordset As ADODB.Recordset

sql = "Select count(Name), Name from AnchorBolt_BM group by Name "

Set local_adoRecordset = adoConnection1.Execute(sql)

If local_adoRecordset.EOF Then Exit Sub

AB_Count = 0
Do While Not local_adoRecordset.EOF
            AB_Count = AB_Count + 1
            local_adoRecordset.MoveNext
Loop
local_adoRecordset.MoveFirst
ReDim AB_BM(1 To AB_Count, 1 To 2) As String
AB_Count = 0
Do While Not local_adoRecordset.EOF
            AB_Count = AB_Count + 1
            AB_BM(AB_Count, 1) = local_adoRecordset!Name
            AB_BM(AB_Count, 2) = local_adoRecordset!expr1000
            local_adoRecordset.MoveNext
Loop
Set local_adoRecordset = Nothing


End Sub
Private Sub ls_HTB_BM(ByVal valPath As String)
Dim ls_String01 As String, ls_String02 As String
Dim i As Integer
Dim intCount As Integer, intCount02 As Integer
Dim strArray() As String, strArray02() As String
Dim strHTB_Size() As String, intHTB As Integer

intHTB = 0
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount = 34 Then
                                    strArray02 = Split(strArray(0), "_")
                                    If strArray02(0) = "HTB" Then
                                                intHTB = intHTB + 1
                                    End If
                        End If
            Loop
Close #1
ReDim strHTB_Size(1 To intHTB) As String
i = 1
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount = 34 Then
                                    strArray02 = Split(strArray(0), "_")
                                    If strArray02(0) = "HTB" Then
                                                strHTB_Size(i) = strArray02(1)
                                                i = i + 1
                                    End If
                        End If
            Loop
Close #1

Dim sql As String
sql = "Delete from HTB_BM"
adoConnection1.Execute (sql)

For i = 1 To intHTB
            sql = "Insert into HTB_BM values ('" & strHTB_Size(i) & "')"
            adoConnection1.Execute (sql)
Next i

Dim local_adoRecordset As ADODB.Recordset


sql = "Select count(Name), Name from HTB_BM group by Name "

Set local_adoRecordset = adoConnection1.Execute(sql)

If local_adoRecordset.EOF Then Exit Sub

HTB_Count = 0
Do While Not local_adoRecordset.EOF
            HTB_Count = HTB_Count + 1
            local_adoRecordset.MoveNext
Loop
local_adoRecordset.MoveFirst
ReDim HTB_BM(1 To HTB_Count, 1 To 2) As String
HTB_Count = 0
Do While Not local_adoRecordset.EOF
            HTB_Count = HTB_Count + 1
            HTB_BM(HTB_Count, 1) = local_adoRecordset!Name
            HTB_BM(HTB_Count, 2) = local_adoRecordset!expr1000
            local_adoRecordset.MoveNext
Loop
Set local_adoRecordset = Nothing
End Sub
Private Sub ls_Plate_BM(ByVal valPath As String)
Dim ls_String01 As String, ls_String02 As String
Dim i As Integer, j As Integer
Dim intCount As Integer, intCount02 As Integer
Dim strArray() As String, strArray02() As String
Dim strPlate_Name() As String, intPlate As Integer
Dim strPlate_Thk() As String, strPlate_Grade() As String, strPlate_Volumn() As String, strPlate_Weight() As String
Dim strPlate_Material() As String

intPlate = 0
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount >= 1 Then
'                        If intCount = 32 Or intCount = 33 Then

                        strArray02 = Split(strArray(0), "_")
                                    If strArray02(0) = "BP" Or strArray02(0) = "EP" Or strArray02(0) = "GP" Or strArray02(0) = "SP" Or strArray02(0) = "RP" Then
                                                intPlate = intPlate + 1
                                    End If
                        End If
            Loop
Close #1

ReDim strPlate_Name(1 To intPlate) As String
ReDim strPlate_Thk(1 To intPlate) As String
ReDim strPlate_Grade(1 To intPlate) As String
ReDim strPlate_Volumn(1 To intPlate) As String
ReDim strPlate_Weight(1 To intPlate) As String
ReDim strPlate_Material(1 To intPlate) As String
i = 1
Open valPath For Input As #1
            Do While Not EOF(1)
                        Input #1, ls_String01
                        strArray = Split(ls_String01)
                        intCount = UBound(strArray, 1)
                        If intCount >= 1 Then
                                    strArray02 = Split(strArray(0), "_")
'                        If intCount = 32 Or intCount = 33 Then
                                    If strArray02(0) = "BP" Or strArray02(0) = "EP" Or strArray02(0) = "GP" Or strArray02(0) = "SP" Or strArray02(0) = "RP" Then
                                                
                                                strPlate_Name(i) = strArray02(0)
                                                strPlate_Thk(i) = strArray02(1)
'                                                bln_Material = False: bln_Grade = False: bln_Volumn = False: bln_Weight = False
                                                For j = 1 To intCount
                                                            If strArray(j) <> "" Then
                                                                        If strPlate_Material(i) = "" Then strPlate_Material(i) = strArray(j): GoTo Label100
                                                                        If strPlate_Material(i) <> "" And strPlate_Grade(i) = "" Then strPlate_Grade(i) = strArray(j): GoTo Label100
                                                                        If strPlate_Material(i) <> "" And strPlate_Grade(i) <> "" And _
                                                                                    strPlate_Volumn(i) = "" Then strPlate_Volumn(i) = strArray(j): GoTo Label100
                                                                        If strPlate_Material(i) <> "" And strPlate_Grade(i) <> "" And _
                                                                                    strPlate_Volumn(i) <> "" And strPlate_Weight(i) = "" Then _
                                                                                    strPlate_Weight(i) = strArray(j): GoTo Label100
                                                            End If
Label100:
                                                Next j
                                                 i = i + 1
                                    End If
                        End If
            Loop
Close #1

Dim sql As String
sql = "Delete from Plate_BM"
adoConnection1.Execute (sql)

For i = 1 To intPlate
            sql = "Insert into Plate_BM values ('" & strPlate_Name(i) & "', '"
            sql = sql & strPlate_Thk(i) & "', '"
            sql = sql & strPlate_Grade(i) & "', '"
            sql = sql & strPlate_Volumn(i) & "', '"
            sql = sql & strPlate_Weight(i) & "')"
            
            adoConnection1.Execute (sql)
Next i

Dim local_adoRecordset As ADODB.Recordset
'Dim HTB_Count As Integer
'
sql = "Select count(Thickness), Thickness from Plate_BM group by Thickness "
Set local_adoRecordset = adoConnection1.Execute(sql)

If local_adoRecordset.EOF Then Exit Sub

Plate_Count = 0
Do While Not local_adoRecordset.EOF
            Plate_Count = Plate_Count + 1
            local_adoRecordset.MoveNext
Loop
local_adoRecordset.MoveFirst
ReDim Plate_BM(1 To Plate_Count, 1 To 3) As String
Plate_Count = 0
Do While Not local_adoRecordset.EOF
            Plate_Count = Plate_Count + 1
            Plate_BM(Plate_Count, 1) = local_adoRecordset!Thickness
'            HTB_BM(HTB_Count, 2) = local_adoRecordset!expr1000
            local_adoRecordset.MoveNext
Loop

Dim Temp As Single, TempWeight As Single, TempVolumn As Single
TempWeight = 0: Temp = 0
For i = 1 To Plate_Count
            sql = "Select Weight from Plate_BM where Thickness = '" & Plate_BM(i, 1) & "'"
            Set local_adoRecordset = adoConnection1.Execute(sql)
            Do While Not local_adoRecordset.EOF
                        Temp = CSng(local_adoRecordset!Weight)
                        local_adoRecordset.MoveNext
                        TempWeight = TempWeight + Temp
            Loop
            Plate_BM(i, 2) = TempWeight
Next i
TempVolumn = 0: Temp = 0
For i = 1 To Plate_Count
            sql = "Select Volumn from Plate_BM where Thickness = '" & Plate_BM(i, 1) & "'"
            Set local_adoRecordset = adoConnection1.Execute(sql)
            Do While Not local_adoRecordset.EOF
                        Temp = CSng(local_adoRecordset!Volumn)
                        local_adoRecordset.MoveNext
                        TempVolumn = TempVolumn + Temp
            Loop
            Plate_BM(i, 3) = TempVolumn
Next i


Set local_adoRecordset = Nothing



End Sub

Private Sub Form_Load()
PBar1.Visible = False
PBar2.Visible = False

End Sub

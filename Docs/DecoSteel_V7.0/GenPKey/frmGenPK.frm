VERSION 5.00
Begin VB.Form frmGenPK 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Generation of Product Key for DecoSteel"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame3 
      Caption         =   "License Key Date "
      Height          =   1155
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   6645
      Begin VB.TextBox txtDuration 
         Height          =   270
         Left            =   2040
         TabIndex        =   9
         Top             =   450
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "개월 동안 사용 가능"
         Height          =   180
         Left            =   4500
         TabIndex        =   10
         Top             =   510
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "사용 기간을 입력 :"
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   510
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdLicense 
      Caption         =   "License file Generation"
      Height          =   1065
      Left            =   6720
      TabIndex        =   6
      Top             =   2010
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   6690
      TabIndex        =   5
      Top             =   3180
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Certification Key"
      Height          =   945
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6645
      Begin VB.TextBox txtPKey 
         BackColor       =   &H00808080&
         Height          =   345
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   6315
      End
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "CK Generation"
      Height          =   1875
      Left            =   6720
      TabIndex        =   2
      Top             =   30
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Serial Key"
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6645
      Begin VB.TextBox txtSerial 
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   330
         Width           =   6435
      End
   End
End
Attribute VB_Name = "frmGenPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGen_Click()
Dim ls_CertiKey As String, valPW As String
Dim String1 As String, String2 As String, String3 As String, String4 As String, String5 As String
txtPKey.Text = ""

valPW = CStr(Trim(txtSerial.Text))
If IsNumeric(valPW) = True Then
             String1 = Format(CDbl(valPW) / 3, "0.00")
             String1 = CStr(Trim(Left(String1, 1))) & CStr(Trim(Mid(String1, 3, 2)))
             String2 = Format(CDbl(valPW) / 13, "0.00")
             String2 = CStr(Trim(Left(String2, 1))) & CStr(Trim(Mid(String2, 3, 3)))
             String3 = Format(CDbl(valPW) / 23, "0.00")
             String3 = CStr(Trim(Left(String3, 1))) & CStr(Trim(Mid(String3, 4, 2)))
             String4 = Format(CDbl(valPW) / 33, "0.00")
             String4 = CStr(Trim(Left(String4, 1))) & CStr(Trim(Mid(String4, 4, 3)))
             String5 = Format(CDbl(valPW) / 43, "0.00")
             String5 = CStr(Trim(Left(String5, 1))) & CStr(Trim(Mid(String5, 3, 2)))
Else
             MsgBox "문자는 지원하지 않습니다."
             Exit Sub
End If

ls_CertiKey = String1 & String2 & String3 & String4 & String5

txtPKey.Text = ls_CertiKey
txtPKey.BackColor = &H808080
End Sub

Private Sub cmdLicense_Click()
Dim MyDate, RecordDate
Dim PathName As String, DateString As String, CurrentDate As String
'Dim S(1 To 8) As String
'Dim StartDate As String, EndDate As String
'Dim StartEA As Integer, EndEA As Integer, i As Integer
'
Dim String1 As String, String2 As String

MyDate = Date

RecordDate = DateAdd("m", Int(Trim(txtDuration.Text)), MyDate)

CurrentDate = pfunReturnKeytoSpace(Date, "-")
DateString = pfunReturnKeytoSpace(RecordDate, "-")

'For i = 1 To 8
'             Randomize
'             S(i) = CStr(Int((9 * Rnd)))
'Next i

'StartEA = Len(CurrentDate): EndEA = Len(DateString)

'For i = 1 To StartEA
'             StartDate = StartDate & S(i) & Mid(CurrentDate, i, 1)
'Next i
'
'For i = 1 To 8
'             Randomize
'             S(i) = CStr(Int((9 * Rnd)))
'Next i
'
'For i = 1 To EndEA
'             EndDate = EndDate & Mid(DateString, i, 1) & S(i)
'Next i

String1 = pfunReturnKeytoSpace(Format(CDbl(CDbl(CurrentDate) / 44), "0.00"), ".")
String2 = pfunReturnKeytoSpace(Format(CDbl(CDbl(DateString) / 48), "0.00"), ".")



 PathName = App.Path & "\License"

Open PathName For Output As #1
             Print #1, String1
             Print #1, String2
Close #1



End Sub

Private Sub Form_Load()

txtPKey.BackColor = &H80000005

End Sub

Private Sub txtSerial_Change()
txtPKey.BackColor = &H80000005
txtPKey.Text = ""

End Sub

Private Function pfunReturnKeytoSpace(ByVal paraString As String, strSource As String) As String
    Dim i               As Integer
    Dim intCount        As Integer
    Dim StrReturnValue  As String
    Dim strArray()      As String
    
    strArray = Split(paraString, strSource) 'Chr(13) & Chr(10) spilt
    intCount = UBound(strArray, 1)                  '인자 갯수 구함
        
    If intCount < 0 Then
        gfunReturnKeytoSpace = ""
        Exit Function
    End If
    
    ReDim ArgBuf(0 To intCount)                     '아규먼트 buffer
    
    StrReturnValue = ""
    
    For i = 0 To intCount
        StrReturnValue = StrReturnValue & strArray(i) '& " "
    Next

    pfunReturnKeytoSpace = Trim(StrReturnValue)
End Function


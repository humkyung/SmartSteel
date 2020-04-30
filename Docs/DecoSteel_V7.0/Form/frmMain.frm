VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00404040&
   Caption         =   "DecoSteel - PML Auto Generation for Connection Detail of Steel Structure"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11655
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   225
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   8655
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSpec 
      Caption         =   "Connection Spec."
      Begin VB.Menu mnuGeneral 
         Caption         =   "General"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDecoMenu 
         Caption         =   "DecoSteel Menu"
      End
      Begin VB.Menu mnuMicro 
         Caption         =   "MicroStation Menu"
      End
      Begin VB.Menu mnuSpace2000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCode 
         Caption         =   "Code Import"
      End
      Begin VB.Menu mnuCodeDrop 
         Caption         =   "Code Drop"
      End
      Begin VB.Menu mnuSpace10000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHBP 
         Caption         =   "Base Plate-Hinged Type"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVBC 
         Caption         =   "Vert Bracing Connection"
      End
      Begin VB.Menu mnuHBC 
         Caption         =   "Hori Bracing Connection"
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMC 
         Caption         =   "Moment Connect (End Plate)"
      End
      Begin VB.Menu mnuSC 
         Caption         =   "Shear Connect"
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuBP 
      Caption         =   "Base Plate"
      Begin VB.Menu mnuHT 
         Caption         =   "Hinged Type"
      End
      Begin VB.Menu mnuSpa01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFT 
         Caption         =   "Fixed Type"
         Begin VB.Menu mnuT01 
            Caption         =   "Type 01"
         End
         Begin VB.Menu mnuSpace21 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT02 
            Caption         =   "Type 02"
         End
         Begin VB.Menu mnuSpace22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT03 
            Caption         =   "Type 03"
         End
         Begin VB.Menu mnuSpace29 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT04 
            Caption         =   "Type 04"
         End
         Begin VB.Menu mnuSpace23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT05 
            Caption         =   "Type 05"
         End
         Begin VB.Menu mnuSpace24 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT06 
            Caption         =   "Type 06"
         End
         Begin VB.Menu mnuSpace25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuT07 
            Caption         =   "Type 07"
         End
      End
   End
   Begin VB.Menu mnuGVB 
      Caption         =   "Vertical Bracing"
      Begin VB.Menu mnuGVBXZ 
         Caption         =   "X-Z Plan"
         Begin VB.Menu mnuVM01 
            Caption         =   "Left-Bottom"
         End
         Begin VB.Menu mnuspace31 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM02 
            Caption         =   "Right-Bottom"
         End
         Begin VB.Menu mnuspace32 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM03 
            Caption         =   "Left-Top"
         End
         Begin VB.Menu mnuspace33 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM04 
            Caption         =   "Right-Top"
         End
         Begin VB.Menu mnuspace34 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM05 
            Caption         =   "Left-Top Offset "
         End
         Begin VB.Menu mnuspace35 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM06 
            Caption         =   "Right-Top Offset"
         End
         Begin VB.Menu mnuspace36 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM07 
            Caption         =   "X-Bracing"
         End
         Begin VB.Menu mnuspace37 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVM08 
            Caption         =   "K-Bracing"
         End
      End
      Begin VB.Menu mnuSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGVBYZ 
         Caption         =   "Y-Z Plan"
         Begin VB.Menu mnuVMY01 
            Caption         =   "Left-Bottom"
         End
         Begin VB.Menu mnuSpace51 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY02 
            Caption         =   "Right-Bottom"
         End
         Begin VB.Menu mnuSpace52 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY03 
            Caption         =   "Left-Top"
         End
         Begin VB.Menu mnuSpace53 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY04 
            Caption         =   "Right-Top"
         End
         Begin VB.Menu mnuSpace54 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY05 
            Caption         =   "Left-Top Offset "
         End
         Begin VB.Menu mnuSpace55 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY06 
            Caption         =   "Right-Top Offset "
         End
         Begin VB.Menu mnuSpace56 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY07 
            Caption         =   "X-Bracing"
         End
         Begin VB.Menu mnuSpace57 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVMY08 
            Caption         =   "K-Bracing"
         End
      End
   End
   Begin VB.Menu mnuGHBXY 
      Caption         =   "Horizontal Bracing"
      Begin VB.Menu mnuHM01 
         Caption         =   "Left_Bottom "
      End
      Begin VB.Menu mnuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM02 
         Caption         =   "Right-Bottom"
      End
      Begin VB.Menu mnuSpace171 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM03 
         Caption         =   "Right-Top"
      End
      Begin VB.Menu mnuSpace7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM04 
         Caption         =   "Left_Top"
      End
      Begin VB.Menu mnuSpace8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM05 
         Caption         =   "X-Bracing"
      End
      Begin VB.Menu mnuSpace9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM06 
         Caption         =   "K-Bracing"
      End
      Begin VB.Menu mnuSpace10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHM07 
         Caption         =   "Bracing for Sub Beam"
      End
   End
   Begin VB.Menu mnuM 
      Caption         =   "Moment Connection (End Plate)"
      Begin VB.Menu mnuMXZ 
         Caption         =   "X-Z Plan"
         Begin VB.Menu mnuMX01 
            Caption         =   "General Location"
            Begin VB.Menu mnuMX0101 
               Caption         =   "Single - X(+) Dir"
            End
            Begin VB.Menu mnuSpace41 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0102 
               Caption         =   "Single - X(-) Dir"
            End
            Begin VB.Menu mnuSpace42 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0103 
               Caption         =   "Double Beam - Left Beam Large"
            End
            Begin VB.Menu mnuSpace43 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0104 
               Caption         =   "Double Beam - Right Beam Large"
            End
            Begin VB.Menu mnuSpace44 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0105 
               Caption         =   "Double Beam - Equal"
            End
         End
         Begin VB.Menu mnuSpace100 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMX02 
            Caption         =   "Top Location"
            Begin VB.Menu mnuMX0201 
               Caption         =   "Single - X(+) Dir"
            End
            Begin VB.Menu mnuSpace61 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0202 
               Caption         =   "Single - X(-) Dir"
            End
            Begin VB.Menu mnuSpace62 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0203 
               Caption         =   "Double Beam - Left Beam Large"
            End
            Begin VB.Menu mnuSpace63 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0204 
               Caption         =   "Double Beam - Right Beam Large"
            End
            Begin VB.Menu mnuSpace64 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMX0205 
               Caption         =   "Double Beam - Equal"
            End
         End
      End
      Begin VB.Menu mnuSpace11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMYZ 
         Caption         =   "Y-Z Plan"
         Begin VB.Menu mnuMY01 
            Caption         =   "General Location"
            Begin VB.Menu mnuMY0101 
               Caption         =   "Single - Y(-) Dir"
            End
            Begin VB.Menu mnuSpace71 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0102 
               Caption         =   "Single - Y(+) Dir"
            End
            Begin VB.Menu mnuSpace72 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0103 
               Caption         =   "Double Beam - Left Beam Large"
            End
            Begin VB.Menu mnuSpace73 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0104 
               Caption         =   "Double Beam - Right Beam Large"
            End
            Begin VB.Menu mnuSpace74 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0105 
               Caption         =   "Double Beam - Equal"
            End
         End
         Begin VB.Menu mnuSpace200 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMY02 
            Caption         =   "Top Location"
            Begin VB.Menu mnuMY0201 
               Caption         =   "Single - Y(-) Dir"
            End
            Begin VB.Menu mnuSpace81 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0202 
               Caption         =   "Single - Y(+) Dir"
            End
            Begin VB.Menu mnuSpace82 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0203 
               Caption         =   "Double Beam - Left Beam Large"
            End
            Begin VB.Menu mnuSpace83 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0204 
               Caption         =   "Double Beam - Right Beam Large"
            End
            Begin VB.Menu mnuSpace84 
               Caption         =   "-"
            End
            Begin VB.Menu mnuMY0205 
               Caption         =   "Double Beam - Equal"
            End
         End
      End
   End
   Begin VB.Menu mnuS 
      Caption         =   "Shear Connection"
      Begin VB.Menu mnuSXZ 
         Caption         =   "XZ Plan"
         Begin VB.Menu mnuSX01 
            Caption         =   "Beam to Column"
            Begin VB.Menu mnuSX0101 
               Caption         =   "X(+) Dir"
            End
            Begin VB.Menu mnuSpace91 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSX0102 
               Caption         =   "X(-) Dir"
            End
         End
         Begin VB.Menu mnuSpace300 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSX02 
            Caption         =   "Beam to Beam"
            Begin VB.Menu mnuSX0201 
               Caption         =   "X(+) Dir"
            End
            Begin VB.Menu mnuSpace92 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSX0202 
               Caption         =   "X(-) Dir"
            End
         End
      End
      Begin VB.Menu mnuSpace12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSYZ 
         Caption         =   "YZ Plan"
         Begin VB.Menu mnuSY01 
            Caption         =   "Beam to Column"
            Begin VB.Menu mnuSY0101 
               Caption         =   "Y(-) Dir"
            End
            Begin VB.Menu mnuSpace93 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSY0102 
               Caption         =   "Y(+) Dir"
            End
         End
         Begin VB.Menu mnuSpace400 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSY02 
            Caption         =   "Beam to Beam"
            Begin VB.Menu mnuSY0201 
               Caption         =   "Y(-) Dir"
            End
            Begin VB.Menu mnuSpace94 
               Caption         =   "-"
            End
            Begin VB.Menu mnuSY0202 
               Caption         =   "Y(+) Dir"
            End
         End
      End
   End
   Begin VB.Menu mnuBM 
      Caption         =   "B/M"
      Begin VB.Menu mnuDetail 
         Caption         =   "Detail B/M"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Private Sub MDIForm_Load()
Dim SerialNum As Long, TotalByte As Long, FreeByte As Long
Dim CheckFlag As Boolean, PW As String, StartFlag As Boolean, CertiKey As String
Dim S1 As String, S2 As String, S3 As String, S4 As String, S5 As String, S6 As String, S7 As String, S8 As String, _
        S9 As String, S10 As String
gin_Int = 0
Call ADO_Connection
                          
'             Call CheckHardSerial(SerialNum, TotalByte, FreeByte)
'
'             If SerialNum < 0 Then SerialNum = -1 * SerialNum
'
'             Call CheckPW_File(PW, SerialNum, CheckFlag)
'
'             If CheckFlag = False Then
'                          Randomize
'                          S1 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S2 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S3 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S4 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S5 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S6 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S7 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S8 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S9 = CStr(Int((9 * Rnd)))
'                          Randomize
'                          S10 = CStr(Int((9 * Rnd)))
'
'                          PW = S1 & S2 & S3 & S4 & S5 & S6 & S7 & S8 & S9 & S10 & "." & CStr(SerialNum)
'                          gs_PW = PW
'                          Call CreatePW_File(PW)
'             Else
'                          gs_PW = PW
'             End If
'
'             Call CheckCertiKey(CertiKey)
'
'             gs_CertiKey = CertiKey
'             If gs_CertiKey = "" Then
'                          frmProductKey.Show 1
'             End If
'
'             Call CertiSystem(PW, StartFlag)
'
'             If StartFlag = False Then
'                     MsgBox "Certification Key is Wrong !!!"
'                     End
'             End If
'             Dim dateFlag As Boolean
'
'             Call datecheck(dateFlag)
'
'             If dateFlag = False Then
'                          MsgBox "License Expire. Please Contact LG E&C."
'                          Kill App.Path & "\license"
'                          End
'             End If

             Me.BackColor = &H404040

             StatusBar1.Panels(1).Width = 20000
             StatusBar1.Panels(1).Text = "" 'gstr_Job

            Dim i
            i = SetWindowPos(frmMain.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            gin_Chk_Flag01 = 1
            
            Call ls_ProgramStart
             
End Sub
Private Sub datecheck(ByRef Flag As Boolean)
Dim pathName As String, i As Integer
Dim Key(1 To 4) As String, strTemp As String
Dim StartDate As String, EndDate As String
Dim string1 As String, string2 As String
Dim string1_1 As String, string1_2 As String
Dim string2_1 As String, string2_2 As String

Dim len1 As Single, len2 As Single

pathName = App.Path & "\License"
'             pathName = Dir$(App.Path & "\License")
i = 1
Open pathName For Input As #1
      Do While Not EOF(1)
            Input #1, strTemp
             Key(i) = strTemp
             i = i + 1
      Loop
Close #1

len1 = Len(Key(1)): len2 = Len(Key(2))


string1_1 = Left(Key(1), len1 - 2)
string1_2 = Mid(Key(1), len1 - 1, 2)
string2_1 = Left(Key(2), len2 - 2)
string2_2 = Mid(Key(2), len2 - 1, 2)

string1 = string1_1 & "." & string1_2
string2 = string2_1 & "." & string2_2

StartDate = Format(CDbl(CDbl(string1) * 44), "0")
EndDate = Format(CDbl(CDbl(string2) * 48), "0")




If CSng(StartDate) <= CSng(Format(Date, "YYYYMMDD")) And CSng(Format(Date, "YYYYMMDD")) <= CSng(EndDate) Then
      Flag = True
Else
      Flag = False
End If

End Sub

Private Sub CheckCertiKey(ByRef refCertiKey As String)
             Dim pathName As String, i As Integer
             Dim Key(1 To 4) As String, strTemp As String
'             pathName = Dir$(App.Path & "\License")
             pathName = App.Path & "\License"

             i = 1
             Open pathName For Input As #1
                          Do While Not EOF(1)
                                              Input #1, strTemp
                                               Key(i) = strTemp
                                               i = i + 1
                          Loop
             Close #1
             
             refCertiKey = Key(4)
End Sub


Private Sub CertiSystem(ByVal valPW As String, ByRef refConfirm_Flag As Boolean)
Dim pathName As String, strTemp As String
Dim ls_CertiKey As String
Dim PWString() As String
Dim TempEA As Single
Dim i As Integer
Dim string1 As String, string2 As String, String3 As String, String4 As String, String5 As String

string1 = Format(CDbl(valPW) / 3, "0.00")
string1 = CStr(Trim(Left(string1, 1))) & CStr(Trim(Mid(string1, 3, 2)))
string2 = Format(CDbl(valPW) / 13, "0.00")
string2 = CStr(Trim(Left(string2, 1))) & CStr(Trim(Mid(string2, 3, 3)))
String3 = Format(CDbl(valPW) / 23, "0.00")
String3 = CStr(Trim(Left(String3, 1))) & CStr(Trim(Mid(String3, 4, 2)))
String4 = Format(CDbl(valPW) / 33, "0.00")
String4 = CStr(Trim(Left(String4, 1))) & CStr(Trim(Mid(String4, 4, 3)))
String5 = Format(CDbl(valPW) / 43, "0.00")
String5 = CStr(Trim(Left(String5, 1))) & CStr(Trim(Mid(String5, 3, 2)))

ls_CertiKey = string1 & string2 & String3 & String4 & String5

If gs_CertiKey = ls_CertiKey Then
             refConfirm_Flag = True
             If gin_Int = 1 Then
                          pathName = App.Path & "\License"
                          Open pathName For Append As #1
                                       Print #1, ls_CertiKey
                          Close #1
             End If
Else
             refConfirm_Flag = False
End If

End Sub

Private Sub CheckPW_File(ByRef valPW As String, ByVal valSerial As Long, ByRef Flag As Boolean)
Dim pathName As String, File_Name As String
Dim strTemp As String, i As Integer
Dim Key(1 To 4) As String
Dim strArray() As String
pathName = App.Path & "\License"
File_Name = Dir$(pathName)

If File_Name = "" Then
             MsgBox "License is not found. Please Contact LG E&C."
             End
Else
             i = 1
             Open pathName For Input As #1
                          Do While Not EOF(1)
                                              Input #1, strTemp
                                               Key(i) = strTemp
                                               i = i + 1
                          Loop
             Close #1
             
             If Key(3) <> "" Then
                        strArray = Split(Key(3), ".")
                        If CStr(valSerial) = Trim(strArray(1)) Then
                                    Flag = True
                                    valPW = Key(3)
                          Else
                                    MsgBox "License is not correct your system. Please Contact LG E&C."
'                                    Kill App.Path & "\license"
                                    End
                          End If
             Else
                          Flag = False
             End If
End If

End Sub
Private Sub CreatePW_File(ByVal valPW As String)
Dim pathName As String

pathName = App.Path & "\License"
Open pathName For Append As #1
    Print #1, valPW
Close #1

End Sub
Private Sub CheckHardSerial(ByRef refSerial As Long, refTbyte As Long, refFbyte As Long)
Dim txt As String
Dim volume_name As String
Dim file_system_name As String
Dim serial_number As Long
Dim component_length As Long
Dim system_flags As Long
Dim sectors_per_cluster As Long
Dim bytes_per_sector As Long
Dim free_clusters As Long
Dim total_clusters As Long
Dim total_bytes As Long
Dim free_bytes As Long

    volume_name = Space$(256)
    file_system_name = Space$(256)
    If GetVolumeInformation("C:\", volume_name, Len(volume_name), serial_number, component_length, _
        system_flags, file_system_name, Len(file_system_name)) = 0 Then
        
        txt = "Error in GetVolumeInformation."
    Else
        refSerial = serial_number
       
    End If

    If GetDiskFreeSpace("C:\", sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters) = 0 Then
        txt = txt & vbCrLf & "Error in GetDiskFreeSpace."
    Else
        
        refTbyte = total_clusters '* sectors_per_cluster) ' * bytes_per_sector) / 10000000000#
        refFbyte = free_clusters '* sectors_per_cluster) '* bytes_per_sector) / 10000000000#
       
    End If

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
'    Dim i As Integer
'For i = CInt(Trim(Me.Height)) To 0 Step -10
'               Me.Width = i
'               Me.Height = i
'Next i

    Call ADO_disConnection
End Sub

Private Sub mnuGVB01_Click()

'frmMoudle_01.Show

End Sub

Private Sub mnuCode_Click()

frmCodeImport.Show

End Sub

Private Sub mnuCodeDrop_Click()
frmTableDrop.Show
End Sub

Private Sub mnuDecoMenu_Click()
'gin_MenuFlag 가 1 : DecoSteel Menu Load, 2 : MicroStation Menu Load
gin_MenuFlag = 1
frmCopyPath.Show


End Sub
Private Sub mnuMicro_Click()
'gin_MenuFlag 가 1 : DecoSteel Menu Load, 2 : MicroStation Menu Load
gin_MenuFlag = 2
frmCopyPath.Show

End Sub

Private Sub mnuDetail_Click()
frmDetailBM.Show
End Sub

Private Sub mnuExit_Click()
'Dim i As Integer
'For i = CInt(Trim(Me.Height)) To 0 Step -10
'               Me.Width = i
'               Me.Height = i
'Next i

Unload Me
End Sub

Private Sub mnuGeneral_Click()
frmGeneral.Show
End Sub



Private Sub mnuHBC_Click()
frmSpecHB.Show
End Sub

Private Sub mnuHBP_Click()
frmSpecHBP.Show
End Sub

Private Sub mnuHM01_Click()
gstr_HBGP_Flag = "Module01"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM02_Click()
gstr_HBGP_Flag = "Module02"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM03_Click()
gstr_HBGP_Flag = "Module03"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM04_Click()
gstr_HBGP_Flag = "Module04"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM05_Click()
gstr_HBGP_Flag = "Module05"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM06_Click()
gstr_HBGP_Flag = "Module06"
Call HBGP_Form_Control
End Sub

Private Sub mnuHM07_Click()
gstr_HBGP_Flag = "Module07"
Call HBGP_Form_Control
End Sub

Private Sub mnuHT_Click()
FrmBasePlate_Hinged.Show
End Sub

Private Sub mnuMC_Click()
    frmSpecMC.Show
End Sub

Private Sub mnuMX0101_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type01"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0102_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type02"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0103_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type03"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0104_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type04"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0105_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type05"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0201_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type01"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0202_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type02"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0203_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type03"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0204_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type04"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMX0205_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type05"
gstr_MCdir_Flag = "X"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0101_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type01"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0102_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type02"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0103_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type03"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0104_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type04"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0105_Click()
gstr_MCEP_Flag = "Module01"
gstr_MCEP_type = "Type05"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0201_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type01"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0202_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type02"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0203_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type03"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0204_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type04"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuMY0205_Click()
gstr_MCEP_Flag = "Module02"
gstr_MCEP_type = "Type05"
gstr_MCdir_Flag = "Y"
Call MOEP_Form_Control
End Sub

Private Sub mnuSC_Click()
    frmSpecSC.Show
End Sub

Private Sub mnuSX0101_Click()
gstr_SCP_Flag = "Module01"
gstr_SCP_type = "Type01"
gstr_SCdir_Flag = "X"
Call SCP_Form_Control
End Sub

Private Sub mnuSX0102_Click()
gstr_SCP_Flag = "Module01"
gstr_SCP_type = "Type02"
gstr_SCdir_Flag = "X"
Call SCP_Form_Control
End Sub

Private Sub mnuSX0201_Click()
gstr_SCP_Flag = "Module02"
gstr_SCP_type = "Type01"
gstr_SCdir_Flag = "X"
Call SCP_Form_Control
End Sub

Private Sub mnuSX0202_Click()
gstr_SCP_Flag = "Module02"
gstr_SCP_type = "Type02"
gstr_SCdir_Flag = "X"
Call SCP_Form_Control
End Sub

Private Sub mnuSY0101_Click()
gstr_SCP_Flag = "Module01"
gstr_SCP_type = "Type01"
gstr_SCdir_Flag = "Y"
Call SCP_Form_Control
End Sub

Private Sub mnuSY0102_Click()
gstr_SCP_Flag = "Module01"
gstr_SCP_type = "Type02"
gstr_SCdir_Flag = "Y"
Call SCP_Form_Control
End Sub

Private Sub mnuSY0201_Click()
gstr_SCP_Flag = "Module02"
gstr_SCP_type = "Type01"
gstr_SCdir_Flag = "Y"
Call SCP_Form_Control
End Sub

Private Sub mnuSY0202_Click()
gstr_SCP_Flag = "Module02"
gstr_SCP_type = "Type02"
gstr_SCdir_Flag = "Y"
Call SCP_Form_Control
End Sub

Private Sub mnuT01_Click()
gstr_BPF_Flag = "Type01"
Call BPF_Form_Control
End Sub

Private Sub mnuT02_Click()
gstr_BPF_Flag = "Type02"
Call BPF_Form_Control

End Sub

Private Sub mnuT03_Click()
gstr_BPF_Flag = "Type03"
Call BPF_Form_Control

End Sub

Private Sub mnuT04_Click()
gstr_BPF_Flag = "Type04"
Call BPF_Form_Control
End Sub

Private Sub mnuT05_Click()
gstr_BPF_Flag = "Type05"
Call BPF_Form_Control
End Sub

Private Sub mnuT06_Click()
gstr_BPF_Flag = "Type06"
Call BPF_Form_Control

End Sub

Private Sub mnuT07_Click()
gstr_BPF_Flag = "Type07"
Call BPF_Form_Control

End Sub

Private Sub mnuVBC_Click()
    frmSpecVB.Show
End Sub

Private Sub BPF_Form_Control()

Select Case gstr_BPF_Flag
    Case "Type01"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 01"
        frmBasePlate_Fixed.Show
    Case "Type02"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 02"
        frmBasePlate_Fixed.Show
    Case "Type03"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 03"
        frmBasePlate_Fixed.Show
    Case "Type04"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 04"
        frmBasePlate_Fixed.Show
    Case "Type05"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 05"
        frmBasePlate_Fixed.Show
    Case "Type06"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 06"
        frmBasePlate_Fixed.Show
    Case "Type07"
        Unload frmBasePlate_Fixed
        frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 07"
        frmBasePlate_Fixed.Show
End Select


End Sub

Private Sub mnuVM01_Click()

gstr_VBGP_Flag = "Module01"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM02_Click()

gstr_VBGP_Flag = "Module02"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM03_Click()

gstr_VBGP_Flag = "Module03"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM04_Click()

gstr_VBGP_Flag = "Module04"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM05_Click()

gstr_VBGP_Flag = "Module05"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM06_Click()

gstr_VBGP_Flag = "Module06"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM07_Click()

gstr_VBGP_Flag = "Module07"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub
Private Sub mnuVM08_Click()

gstr_VBGP_Flag = "Module08"
gstr_VBdir_Flag = "X"
Call VBGP_Form_Control

End Sub

Private Sub MOEP_Form_Control()
Select Case gstr_MCEP_Flag
    Case "Module01"
        Unload frmMoudle_Mc
        If gstr_MCdir_Flag = "Y" Then
            frmMoudle_Mc.Caption = "Module 01 for End Plate of Moment Connection for YZ Plane"
        Else
            frmMoudle_Mc.Caption = "Module 01 for End Plate of Moment Connection for XZ Plane"
        End If
        frmMoudle_Mc.Show
    Case "Module02"
        Unload frmMoudle_Mc
        If gstr_MCdir_Flag = "Y" Then
            frmMoudle_Mc.Caption = "Module 02 for End Plate of Moment Connection for YZ Plane"
        Else
            frmMoudle_Mc.Caption = "Module 02 for End Plate of Moment Connection for XZ Plane"
        End If
        frmMoudle_Mc.Show
End Select

End Sub

Private Sub SCP_Form_Control()
Select Case gstr_SCP_Flag
    Case "Module01"
        Unload frmMoudle_Sc
        If gstr_SCdir_Flag = "Y" Then
            frmMoudle_Sc.Caption = "Module 01 for Shear Connection for YZ Plane"
        Else
            frmMoudle_Sc.Caption = "Module 01 for Shear Connection for XZ Plane"
        End If
        frmMoudle_Sc.Show
    Case "Module02"
        Unload frmMoudle_Sc
        If gstr_SCdir_Flag = "Y" Then
            frmMoudle_Sc.Caption = "Module 02 for Shear Connection for YZ Plane"
        Else
            frmMoudle_Sc.Caption = "Module 02 for Shear Connection for XZ Plane"
        End If
        frmMoudle_Sc.Show
End Select

End Sub

Private Sub VBGP_Form_Control()

Select Case gstr_VBGP_Flag
    Case "Module01"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 01 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 01 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 01 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 01 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module02"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 02 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 02 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 02 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 02 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module03"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 03 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 03 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 03 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 03 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module04"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 04 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 04 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 04 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 04 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module05"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 05 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 05 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 05 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 05 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module06"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 06 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 06 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 06 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 06 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module07"
        Unload frmMoudle_Ver
        If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 07 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 07 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 07 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 07 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
        frmMoudle_Ver.Show
    Case "Module08"
       Unload frmMoudle_Ver
       If gstr_VBdir_Flag = "Y" Then
            'frmMoudle_Ver.Caption = "Module 08 for Gusset Plate of Vertical Bracing for YZ Plane"
            gs_Caption = "Module 08 for Gusset Plate of Vertical Bracing for YZ Plane"
        Else
            'frmMoudle_Ver.Caption = "Module 08 for Gusset Plate of Vertical Bracing for XZ Plane"
            gs_Caption = "Module 08 for Gusset Plate of Vertical Bracing for XZ Plane"
        End If
       frmMoudle_Ver.Show
        
End Select

End Sub

Private Sub HBGP_Form_Control()

Select Case gstr_HBGP_Flag
    Case "Module01"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 01 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 01 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module02"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 02 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 02 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module03"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 03 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 03 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module04"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 04 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 04 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module05"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 05 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 05 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module06"
        Unload frmMoudle_Hor
        'frmMoudle_Hor.Caption = "Module 06 for Gusset Plate of Horizontal Bracing"
        gs_Caption = "Module 06 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
    Case "Module07"
        Unload frmMoudle_Hor
        gs_Caption = "Module 07 for Gusset Plate of Horizontal Bracing"
        'frmMoudle_Hor.Caption = "Module 07 for Gusset Plate of Horizontal Bracing"
        frmMoudle_Hor.Show
End Select

End Sub

Private Sub mnuVMY01_Click()
gstr_VBGP_Flag = "Module01"
gstr_VBdir_Flag = "Y"
Call VBGP_Form_Control
End Sub

Private Sub mnuVMY02_Click()
gstr_VBGP_Flag = "Module02"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY03_Click()
gstr_VBGP_Flag = "Module03"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY04_Click()
gstr_VBGP_Flag = "Module04"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY05_Click()
gstr_VBGP_Flag = "Module05"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY06_Click()
gstr_VBGP_Flag = "Module06"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY07_Click()
gstr_VBGP_Flag = "Module07"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub

Private Sub mnuVMY08_Click()
gstr_VBGP_Flag = "Module08"
gstr_VBdir_Flag = "Y"

Call VBGP_Form_Control
End Sub
Private Sub ls_ProgramStart()
 If Len(Command) = 0 Then
           'frmGeneral.Show
      MsgBox "Please Start in FrameWorks Plus...."
      End
Else
           Me.WindowState = vbMinimized
           Select Case CStr(Trim(Command))
                       Case "10"
                              frmGeneral.Show
                       Case "20"
                              frmCodeImport.Show
                       Case "30"
                              frmTableDrop.Show
                       Case "40"
                              frmSpecHBP.Show
                       Case "50"
                              frmSpecVB.Show
                       Case "60"
                              frmSpecHB.Show
                       Case "70"
                              frmSpecMC.Show
                       Case "80"
                              frmSpecSC.Show
                       Case "100"  ' Hinged Base Plate Call
                                   FrmBasePlate_Hinged.Show
                       Case "201" ' Fixed Base Plate type 01
                                  gstr_BPF_Flag = "Type01"
                                  frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 01"
                                  frmBasePlate_Fixed.Show
                       Case "202" ' Fixed Base Plate type 02
                                   gstr_BPF_Flag = "Type02"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 02"
                                   frmBasePlate_Fixed.Show
                       Case "203" ' Fixed Base Plate type 03
                                   gstr_BPF_Flag = "Type03"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 03"
                                   frmBasePlate_Fixed.Show
                       Case "204" ' Fixed Base Plate type 04
                                   gstr_BPF_Flag = "Type04"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 04"
                                   frmBasePlate_Fixed.Show
                       Case "205" ' Fixed Base Plate type 05
                                   gstr_BPF_Flag = "Type05"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 05"
                                   frmBasePlate_Fixed.Show
                       Case "206" ' Fixed Base Plate type 06
                                   gstr_BPF_Flag = "Type06"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 06"
                                   frmBasePlate_Fixed.Show
                       Case "207" ' Fixed Base Plate type 07
                                   gstr_BPF_Flag = "Type07"
                                   frmBasePlate_Fixed.Caption = "Base Plate Fixed Type 07"
                                   frmBasePlate_Fixed.Show
                       Case "301" 'Vertical Gusset Plate XZ Plan Left-Bottom
                                   gstr_VBGP_Flag = "Module01"
                                   gstr_VBdir_Flag = "X"
                                  Call VBGP_Form_Control

                       Case "302" 'Vertical Gusset Plate XZ Plan Right-Bottom
                                   gstr_VBGP_Flag = "Module02"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "303" 'Vertical Gusset Plate XZ Plan Left-Top
                                   gstr_VBGP_Flag = "Module03"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "304" 'Vertical Gusset Plate XZ Plan Right-Bottom
                                   gstr_VBGP_Flag = "Module04"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "305" 'Vertical Gusset Plate XZ Plan Left-Top Offset
                                   gstr_VBGP_Flag = "Module05"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "306" 'Vertical Gusset Plate XZ Plan Right-Top Offset
                                   gstr_VBGP_Flag = "Module06"
                                   gstr_VBdir_Flag = "X"
                                  Call VBGP_Form_Control

                       Case "307" 'Vertical Gusset Plate XZ Plan X Bracing
                                   gstr_VBGP_Flag = "Module07"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "308" 'Vertical Gusset Plate XZ Plan K Bracing
                                   gstr_VBGP_Flag = "Module08"
                                   gstr_VBdir_Flag = "X"
                                   Call VBGP_Form_Control

                       Case "311" 'Vertical Gusset Plate YZ Plan Left-Bottom
                                   gstr_VBGP_Flag = "Module01"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "312" 'Vertical Gusset Plate YZ Plan Right-Bottom
                                   gstr_VBGP_Flag = "Module02"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "313" 'Vertical Gusset Plate YZ Plan Left-Top
                                   gstr_VBGP_Flag = "Module03"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "314" 'Vertical Gusset Plate YZ Plan Right-Bottom
                                   gstr_VBGP_Flag = "Module04"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "315" 'Vertical Gusset Plate YZ Plan Left-Top Offset
                                   gstr_VBGP_Flag = "Module05"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "316" 'Vertical Gusset Plate YZ Plan Right-Top Offset
                                   gstr_VBGP_Flag = "Module06"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "317" 'Vertical Gusset Plate YZ Plan X Bracing
                                   gstr_VBGP_Flag = "Module07"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "318" 'Vertical Gusset Plate YZ Plan K Bracing
                                   gstr_VBGP_Flag = "Module08"
                                   gstr_VBdir_Flag = "Y"
                                  Call VBGP_Form_Control
                       Case "401" 'Hori Gusset Plate Left-Bottom
                                   gstr_HBGP_Flag = "Module01"
                                   Call HBGP_Form_Control
                       Case "402" 'Hori Gusset Plate Right-Bottom
                                   gstr_HBGP_Flag = "Module02"
                                   Call HBGP_Form_Control
                       Case "403" 'Hori Gusset Plate Right-Top
                                   gstr_HBGP_Flag = "Module03"
                                   Call HBGP_Form_Control
                       Case "404" 'Hori Gusset Plate Left-Top
                                   gstr_HBGP_Flag = "Module04"
                                   Call HBGP_Form_Control
                       Case "405" 'Hori Gusset Plate X Bracing
                                   gstr_HBGP_Flag = "Module05"
                                   Call HBGP_Form_Control
                       Case "406" 'Hori Gusset Plate K-Bracing
                                   gstr_HBGP_Flag = "Module06"
                                   Call HBGP_Form_Control
                       Case "407" 'Hori Gusset Plate Bracing for Sub Beam
                                   gstr_HBGP_Flag = "Module07"
                                   Call HBGP_Form_Control
                       Case "501"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type01"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "502"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type02"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "503"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type03"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "504"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type04"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "505"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type05"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "506"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type01"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "507"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type02"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "508"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type03"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "509"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type04"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "510"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type05"
                                   gstr_MCdir_Flag = "X"
                                   Call MOEP_Form_Control
                       Case "511"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type02"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "512"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type01"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "513"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type03"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "514"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type04"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "515"
                                   gstr_MCEP_Flag = "Module01"
                                   gstr_MCEP_type = "Type05"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "516"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type02"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "517"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type01"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "518"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type03"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "519"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type04"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "520"
                                   gstr_MCEP_Flag = "Module02"
                                   gstr_MCEP_type = "Type05"
                                   gstr_MCdir_Flag = "Y"
                                   Call MOEP_Form_Control
                       Case "601"
                                   gstr_SCP_Flag = "Module01"
                                   gstr_SCP_type = "Type01"
                                   gstr_SCdir_Flag = "X"
                                   Call SCP_Form_Control
                       Case "602"
                                   gstr_SCP_Flag = "Module01"
                                   gstr_SCP_type = "Type02"
                                   gstr_SCdir_Flag = "X"
                                   Call SCP_Form_Control
                       Case "603"
                                   gstr_SCP_Flag = "Module02"
                                   gstr_SCP_type = "Type01"
                                   gstr_SCdir_Flag = "X"
                                   Call SCP_Form_Control
                       Case "604"
                                   gstr_SCP_Flag = "Module02"
                                   gstr_SCP_type = "Type02"
                                   gstr_SCdir_Flag = "X"
                                   Call SCP_Form_Control
                       Case "611"
                                   gstr_SCP_Flag = "Module01"
                                   gstr_SCP_type = "Type01"
                                   gstr_SCdir_Flag = "Y"
                                   Call SCP_Form_Control
                       Case "612"
                                   gstr_SCP_Flag = "Module01"
                                   gstr_SCP_type = "Type02"
                                   gstr_SCdir_Flag = "Y"
                                   Call SCP_Form_Control
                       Case "613"
                                   gstr_SCP_Flag = "Module02"
                                   gstr_SCP_type = "Type01"
                                   gstr_SCdir_Flag = "Y"
                                   Call SCP_Form_Control
                       Case "614"
                                   gstr_SCP_Flag = "Module02"
                                   gstr_SCP_type = "Type02"
                                   gstr_SCdir_Flag = "Y"
                                   Call SCP_Form_Control
                        Case "700"
                                    gin_MenuFlag = 1
                                    frmCopyPath.Show
                        Case "701"
                                    gin_MenuFlag = 2
                                    frmCopyPath.Show
                        Case "decosteel"
                              Me.WindowState = vbNormal
                              frmGeneral.Show
                        Case Else
                              MsgBox "Please Start in FrameWorks Plus...."
                              
                              Kill gs_FWP & "\MSVCRT5B.lic"
                              Kill gs_LGFWP & "\LGFWP_License.ini"
                              End
           End Select
End If
End Sub

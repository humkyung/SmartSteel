VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMoudle_Ver 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmModule_01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10050
   Begin VB.Frame Frame6 
      Height          =   1035
      Left            =   2190
      TabIndex        =   42
      Top             =   6030
      Width           =   1155
      Begin VB.CheckBox chkNut 
         Caption         =   "Nut Model Check"
         Height          =   615
         Left            =   150
         TabIndex        =   43
         Top             =   240
         Value           =   1  '확인
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selcet Member Code"
      Height          =   585
      Left            =   6720
      TabIndex        =   38
      Top             =   60
      Width           =   3285
      Begin VB.ComboBox cmbCode 
         Height          =   300
         Left            =   1200
         TabIndex        =   39
         Text            =   "JIS"
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmbMake 
      Caption         =   "Make PML"
      Height          =   435
      Left            =   6840
      TabIndex        =   35
      Top             =   6600
      Width           =   1530
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select PML Unit"
      Height          =   1050
      Left            =   60
      TabIndex        =   30
      Top             =   6030
      Width           =   2085
      Begin VB.OptionButton optFeet 
         Caption         =   "feet"
         Height          =   240
         Left            =   1305
         TabIndex        =   34
         Top             =   600
         Width           =   600
      End
      Begin VB.OptionButton optInch 
         Caption         =   "inch"
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   600
         Width           =   690
      End
      Begin VB.OptionButton optM 
         Caption         =   "m"
         Height          =   195
         Left            =   1305
         TabIndex        =   32
         Top             =   330
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optMM 
         Caption         =   "mm"
         Height          =   240
         Left            =   360
         TabIndex        =   31
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Select Bracing Shape"
      Height          =   1710
      Left            =   6705
      TabIndex        =   19
      Top             =   4770
      Width           =   3300
      Begin MSComctlLib.ImageList ImageIcon 
         Left            =   2640
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":0326
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":0640
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":095A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":0C74
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":0F8E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton opt06 
         Caption         =   "Option6"
         Height          =   195
         Left            =   2790
         TabIndex        =   25
         Top             =   555
         Width           =   240
      End
      Begin VB.OptionButton opt05 
         Caption         =   "Option5"
         Height          =   195
         Left            =   2295
         TabIndex        =   24
         Top             =   555
         Width           =   240
      End
      Begin VB.OptionButton opt04 
         Caption         =   "Option4"
         Height          =   195
         Left            =   1755
         TabIndex        =   23
         Top             =   555
         Width           =   240
      End
      Begin VB.OptionButton opt03 
         Caption         =   "Option3"
         Height          =   240
         Left            =   1260
         TabIndex        =   22
         Top             =   555
         Width           =   240
      End
      Begin VB.OptionButton opt02 
         Caption         =   "Option2"
         Height          =   240
         Left            =   720
         TabIndex        =   21
         Top             =   555
         Width           =   240
      End
      Begin VB.OptionButton opt01 
         Caption         =   "Option1"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   555
         Width           =   195
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   2745
         Top             =   885
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   2250
         Top             =   885
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   1710
         Top             =   885
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1215
         Top             =   885
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   675
         Top             =   885
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   135
         Top             =   885
         Width           =   480
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Select Modeling Option"
      Height          =   1200
      Left            =   45
      TabIndex        =   18
      Top             =   4770
      Width           =   3300
      Begin VB.ComboBox cmbType 
         Height          =   300
         Left            =   1950
         TabIndex        =   29
         Text            =   "Type-01"
         Top             =   720
         Width           =   1275
      End
      Begin VB.ComboBox cmbColShape 
         Height          =   300
         Left            =   1560
         TabIndex        =   27
         Text            =   "Strong Axis"
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type of Module-01 :"
         Height          =   180
         Left            =   75
         TabIndex        =   28
         Top             =   795
         Width           =   1725
      End
      Begin VB.Label lblColShape 
         AutoSize        =   -1  'True
         Caption         =   "Column Shape :"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "View of Modeling Option"
      Height          =   2310
      Left            =   3375
      TabIndex        =   17
      Top             =   4770
      Width           =   3300
      Begin VB.TextBox txtSpace 
         Height          =   270
         Left            =   2190
         TabIndex        =   41
         Text            =   "0"
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label lblSpace 
         AutoSize        =   -1  'True
         Caption         =   "Space :"
         Height          =   180
         Left            =   2190
         TabIndex        =   40
         Top             =   930
         Width           =   660
      End
      Begin VB.Image imgModel 
         Height          =   2070
         Left            =   45
         Top             =   180
         Width           =   3180
      End
   End
   Begin VB.Frame fraModule 
      Caption         =   "View of Module-01"
      Height          =   2310
      Left            =   3375
      TabIndex        =   16
      Top             =   45
      Width           =   3300
      Begin VB.Image imgModule 
         Height          =   1950
         Left            =   315
         Top             =   270
         Width           =   2760
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "View Axis"
      Height          =   2310
      Left            =   45
      TabIndex        =   15
      Top             =   45
      Width           =   3300
      Begin VB.Image imgAxis 
         Height          =   1905
         Left            =   225
         Top             =   270
         Width           =   2835
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Your Selection"
      Height          =   1590
      Left            =   6705
      TabIndex        =   8
      Top             =   765
      Width           =   3300
      Begin VB.Label lblBracing 
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
         TabIndex        =   14
         Top             =   1260
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
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   255
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   390
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bracing Size :"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Beam Size   :"
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
         TabIndex        =   10
         Top             =   645
         Width           =   1440
      End
      Begin VB.Label Label1 
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
         TabIndex        =   9
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   8550
      TabIndex        =   6
      Top             =   6585
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select Bracing"
      Height          =   2310
      Left            =   6705
      TabIndex        =   4
      Top             =   2430
      Width           =   3300
      Begin VB.ListBox lstBType 
         Height          =   780
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   3165
      End
      Begin VB.ListBox lstBracing 
         Columns         =   2
         Height          =   1140
         Left            =   90
         TabIndex        =   5
         Top             =   1080
         Width           =   3165
      End
   End
   Begin VB.Frame fraBeam 
      Caption         =   "Select Beam"
      Height          =   2310
      Left            =   3375
      TabIndex        =   1
      Top             =   2430
      Width           =   3300
      Begin VB.ComboBox cmbBeam 
         Height          =   300
         Left            =   90
         TabIndex        =   37
         Top             =   270
         Width           =   3165
      End
      Begin VB.ListBox lstBeam 
         Columns         =   2
         Height          =   1500
         Left            =   90
         TabIndex        =   3
         Top             =   630
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
      Begin MSComctlLib.ImageList ImageM8 
         Left            =   2520
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":12A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":169D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2C0FC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM7 
         Left            =   1920
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   139
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":41826
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":571CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":6CB72
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM6 
         Left            =   1320
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   139
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":82518
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":97EBE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM5 
         Left            =   720
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":AD864
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":C2F8E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM4 
         Left            =   120
         Top             =   1260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   139
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":D86B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":EE05E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":103A04
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1193AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":12ED50
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":14447A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":159BA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":16F2CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1849F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":19A122
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1AF84C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM3 
         Left            =   2520
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   139
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1C4F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1DA91C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":1F02C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":205C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":21B60E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":230D38
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":246462
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":25BB8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2712B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2869E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":29C10A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM2 
         Left            =   1920
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2B1834
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2C6F5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2DC688
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":2F1DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3074DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":31CC06
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":332330
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageM1 
         Left            =   1320
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   212
         ImageHeight     =   138
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":347CD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":35D400
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":372DA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3884D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":39DE76
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3B35A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3C8F46
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   720
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   184
         ImageHeight     =   130
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3DE8EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":3F018E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":401A30
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":4132D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":424B74
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":4361EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":447868
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":45910A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   189
         ImageHeight     =   127
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":46A9AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmModule_01.frx":47C3C6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbColumn 
         Height          =   300
         Left            =   90
         TabIndex        =   36
         Top             =   270
         Width           =   3165
      End
      Begin VB.ListBox lstColumn 
         Columns         =   2
         Height          =   1500
         Left            =   90
         TabIndex        =   2
         Top             =   630
         Width           =   3165
      End
   End
End
Attribute VB_Name = "frmMoudle_Ver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spec_Code As String
Dim gstr_Job As String, gs_PMLunit As String

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

Private Sub cmbCode_Click()

Dim form_Code As String
Dim xSQL As String

form_Code = CStr(Trim(cmbCode.Text))
    

If form_Code = "JIS" Then
      
      cmbBeam.Enabled = True
      cmbColumn.Enabled = True
      
       xSQL = "Select member_sort from code_" & form_Code & " " & _
      "where member_type = 'Hbeam' " & _
      "group by member_sort " & _
      "order by member_sort "
      
      Call Query_AddList2_function(0, cmbColumn, xSQL)
      'Call cmbColumn_Click
      Call Query_AddList2_function(0, cmbBeam, xSQL)
      'Call cmbBeam_Click
      xSQL = "Select member_name from code_" & form_Code & " " & _
      "where member_type = 'hbeam' " & _
      "order by member_no "
Else
       xSQL = "Select member_name from code_" & form_Code & " " & _
      "where member_type = 'hbeam' "
       cmbBeam.Enabled = False
       cmbColumn.Enabled = False
               
End If
Call Query_AddList2_function(0, lstBeam, xSQL)
Call Query_AddList2_function(0, lstColumn, xSQL)


End Sub

Private Sub cmbColShape_click()

Select Case gstr_VBGP_Flag
   Case "Module01", "Module02"
    If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
    Else
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
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
        cmbType.AddItem "Type-07"
    Else
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
    End If
End Select

Call Type_BMP_Control_Start
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
Dim ColumnName As String
Dim BeamName As String
Dim BracingName As String
Dim tempColShape As String, tempType As String, TempUnit As String, TempPath As String
Dim strTemp As String, strType As String, tempSpace3 As Single
Dim TempNut As Integer
Dim formCode As String

On Error GoTo Labelstop

    ColumnName = lblColumn.Caption
    BeamName = lblBeam.Caption
    BracingName = lblBracing.Caption
    tempSpace3 = CSng(Trim(txtSpace.Text))
    strTemp = Trim(lstBType.List(lstBType.ListIndex))
    TempNut = CInt(Trim(chkNut.Value))
    formCode = CStr(Trim(cmbCode.Text))
    
    tempType = cmbType.Text
    tempColShape = cmbColShape.Text
    
    If formCode = "" Then MsgBox "Code Selection Error. You must select Code Selection !!!": Exit Sub
    
    Select Case strTemp
               Case "Angle"
                              strTemp = "A"
               Case "Channel"
                              strTemp = "C"
               Case "Double Angle"
                              strTemp = "DA"
               Case "Double Channel"
                              strTemp = "DC"
               Case "Tee"
                              strTemp = "T"
    End Select
    
    If opt01.Value = True Then
               strType = "opt01"
    ElseIf opt02.Value = True Then
               strType = "opt02"
    ElseIf opt03.Value = True Then
               strType = "opt03"
    ElseIf opt04.Value = True Then
               strType = "opt04"
    ElseIf opt05.Value = True Then
               strType = "opt05"
    ElseIf opt06.Value = True Then
               strType = "opt06"
    End If
    
    Select Case gstr_VBGP_Flag
        Case "Module01", "Module02"
'            If tempType <> "Type-04" Then
                  If ColumnName = "N/A" Then
                      MsgBox "Column을 선택 하십시요."
                      Exit Sub
                  End If
'            End If
        Case "Module03", "Module04"
            If ColumnName = "N/A" Then
                MsgBox "Column을 선택 하십시요."
                Exit Sub
            End If
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
        Case "Module05", "Module06", "Module08"
            If BeamName = "N/A" Then
                MsgBox "Beam을 선택 하십시요."
                Exit Sub
            End If
        Case "Module07"
               If CSng(Trim(txtSpace.Text)) <= 0 Then
                              MsgBox "Space Value is Zero or Minus. Chage Space Value"
                Exit Sub
            End If
    End Select
    If BracingName = "N/A" Then
        MsgBox "Bracing을 선택 하십시요."
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
      Open App.Path & "\Library\" & gstr_VBGP_Flag & "_" & gstr_VBdir_Flag & "_Vert_history.ini" For Output As #1
            Print #1, gstr_Job
            Print #1, Trim(cmbCode.Text)
            Print #1, Trim(lstColumn.ListIndex)
            Print #1, Trim(lstBeam.ListIndex)
            Print #1, Trim(lstBType.ListIndex)
            Print #1, Trim(lstBracing.ListIndex)
      Close #1

'    If ColumnName = "" Then
'        MsgBox "Select Column Member Size....."
'    Else
'        frmMain.CommonDialog.CancelError = True
        frmMain.CommonDialog.InitDir = App.Path
        frmMain.CommonDialog.DialogTitle = "Save PML File "
        frmMain.CommonDialog.Filter = "BasePlate (*.pml)|*.pml|"
        frmMain.CommonDialog.FileName = "Test.pml"
        
        frmMain.CommonDialog.ShowSave
        
        TempPath = frmMain.CommonDialog.FileName
    
        If gstr_VBdir_Flag = "X" Then
            Call VB_PML_XZ(TempPath, gstr_Job, spec_Code, formCode, gstr_VBGP_Flag, tempColShape, tempType, TempUnit, _
                                           ColumnName, BeamName, BracingName, strTemp, strType, tempSpace3, TempNut)
        ElseIf gstr_VBdir_Flag = "Y" Then
            Call VB_PML_YZ(TempPath, gstr_Job, spec_Code, formCode, gstr_VBGP_Flag, tempColShape, tempType, TempUnit, _
                                           ColumnName, BeamName, BracingName, strTemp, strType, tempSpace3, TempNut)
        End If
        Call PML_Run(TempPath)
        End
'    End If

Labelstop:
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
      Me.Caption = "Current Project : " & gstr_Job & ", " & gs_Caption
      
      Open App.Path & "\" & gstr_Job & "_pmlunit.ini" For Input As #1
            Input #1, gs_PMLunit
      Close #1
End If

lstBType.AddItem "Angle"
lstBType.AddItem "Channel"
lstBType.AddItem "Double Angle"
lstBType.AddItem "Double Channel"
lstBType.AddItem "Tee"

cmbColShape.AddItem "Strong Axis"
cmbColShape.AddItem "Weak Axis"


Call gs_CobAddItem(cmbCode)

Dim pre_Code As String, pre_ColumnIndex As String, pre_BeamIndex As String, pre_BtypeIndex As String, pre_BraceIndex As String, _
    pre_Job As String
On Error GoTo Error100
Open App.Path & "\Library\" & gstr_VBGP_Flag & "_" & gstr_VBdir_Flag & "_vert_history.ini" For Input As #1
      Input #1, pre_Job
      Input #1, pre_Code
      Input #1, pre_ColumnIndex
      Input #1, pre_BeamIndex
      Input #1, pre_BtypeIndex
      Input #1, pre_BraceIndex
Close #1

restart01:

cmbCode.Text = pre_Code
Call cmbCode_Click
If pre_Code = "JIS" Then
      sql = "Select member_sort from code_" & pre_Code & " " & _
            "where member_type = 'Hbeam' " & _
            "group by member_sort " & _
            "order by member_sort "

      Call Query_AddList2_function(0, cmbColumn, sql)
      Call cmbColumn_Click
      Call Query_AddList2_function(0, cmbBeam, sql)
      Call cmbBeam_Click
End If

If gstr_Job = pre_Job Then
      If pre_ColumnIndex <> "-1" And pre_ColumnIndex <> "" Then
            lstColumn.Selected(pre_ColumnIndex) = True
      End If
      If pre_BeamIndex <> "-1" And pre_BeamIndex <> "" Then
            lstBeam.Selected(pre_BeamIndex) = True
      End If
      If pre_BtypeIndex <> "-1" And pre_BtypeIndex <> "" Then
            lstBType.Selected(pre_BtypeIndex) = True
      End If
      If pre_BraceIndex <> "-1" And pre_BraceIndex <> "" Then
            lstBracing.Selected(pre_BraceIndex) = True
      End If
End If

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               End
'Else
'               sql = "Select Member_Name from VB_Connection " & _
'                     "where Shape = 'Angle' and job = '" & gstr_Job & "'"
'
'               Call Query_AddList_function(1, lstBracing, sql)
End If

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



Exit Sub

Error100:
Select Case Err.Number

   Case 53   ' "txt파일이 없는 경우의 error 번호(path에 따라 다른 것 같음)
      pre_Code = "JIS"
      GoTo restart01
   Case Else
      ' 여기서 다른 상황을 다룹니다.
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Len(Command) = 0 Or Command = "decosteel" Then
            Unload Me
Else
            End
End If
End Sub

Private Sub lstBType_Click()
Dim strTemp As String
Dim sql As String

If gstr_Job = "" Then
               MsgBox "Job is not selected. Please Job Select at General"
               Exit Sub
Else
               strTemp = Trim(lstBType.List(lstBType.ListIndex))
               sql = "Select Member_Name from VB_Connection" & _
                     " where Shape = '" & strTemp & "' and job = '" & gstr_Job & "'"
               
               Call Query_AddList_function(1, lstBracing, sql)
End If
opt01.Enabled = True
opt02.Enabled = True
opt03.Enabled = True
opt04.Enabled = True
opt05.Enabled = True
opt06.Enabled = True

Select Case strTemp
    Case "Angle"
        opt01.Value = True
        opt03.Enabled = False
        opt04.Enabled = False
        opt05.Enabled = False
        opt06.Enabled = False
    Case "Channel"
        opt01.Enabled = False
        opt02.Enabled = False
        opt03.Value = True
'        opt04.Enabled = False
        opt05.Enabled = False
        opt06.Enabled = False
    Case "Double Angle"
        opt01.Value = True
        opt02.Enabled = False
        opt03.Enabled = False
        opt04.Enabled = False
'        opt05.Enabled = False
'        opt06.Enabled = False
    Case "Double Channel"
        opt01.Enabled = False
        opt02.Enabled = False
        opt03.Value = True
        opt04.Enabled = False
        opt05.Enabled = False
        opt06.Enabled = False

    Case "Tee"
        opt06.Value = True
        opt02.Enabled = False
        opt03.Enabled = False
        opt04.Enabled = False
'        opt05.Enabled = False
'        opt06.Enabled = False

End Select

End Sub

Private Sub lstColumn_Click()

lblColumn.Caption = Trim(lstColumn.List(lstColumn.ListIndex))

End Sub
Private Sub lstBeam_Click()
lblBeam.Caption = Trim(lstBeam.List(lstBeam.ListIndex))

End Sub
Private Sub lstBracing_Click()
Dim xSQL As String
Dim lstr_MemberName As String

Dim retData As ADODB.Recordset

lblBracing.Caption = Trim(lstBracing.List(lstBracing.ListIndex))

lstr_MemberName = CStr(Trim(lstBracing.List(lstBracing.ListIndex)))

xSQL = "Select Code from VB_Connection where Job = '" & gstr_Job & "' "
xSQL = xSQL & "and Member_Name = '" & lstr_MemberName & "'"

Set retData = adoConnection1.Execute(xSQL)

spec_Code = CStr(Trim(retData!code))

retData.Close

Set retData = Nothing
End Sub

Private Sub Form_Control()
opt01.Value = True
opt03.Enabled = False
opt04.Enabled = False
opt05.Enabled = False
opt06.Enabled = False

Set Image3.Picture = ImageIcon.ListImages(1).Picture
Set Image4.Picture = ImageIcon.ListImages(2).Picture
Set Image5.Picture = ImageIcon.ListImages(3).Picture
Set Image6.Picture = ImageIcon.ListImages(4).Picture
Set Image7.Picture = ImageIcon.ListImages(5).Picture
Set Image8.Picture = ImageIcon.ListImages(6).Picture

Select Case gstr_VBGP_Flag
    Case "Module01"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul01.bmp")
        Set imgModule.Picture = ImageList2.ListImages(1).Picture
        fraBeam.Enabled = False
        lstBeam.Visible = False
        cmbBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        fraModule.Caption = "View of Module-01"
    Case "Module02"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul02.bmp")
        Set imgModule.Picture = ImageList2.ListImages(2).Picture
        fraBeam.Enabled = False
        lstBeam.Visible = False
        cmbBeam.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        fraModule.Caption = "View of Module-02"
    Case "Module03"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul03.bmp")
        Set imgModule.Picture = ImageList2.ListImages(3).Picture
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.AddItem "Type-07"
        
        fraModule.Caption = "View of Module-03"
    Case "Module04"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul04.bmp")
        Set imgModule.Picture = ImageList2.ListImages(4).Picture
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        cmbType.AddItem "Type-04"
        cmbType.AddItem "Type-05"
        cmbType.AddItem "Type-06"
        cmbType.AddItem "Type-07"
        fraModule.Caption = "View of Module-04"
    Case "Module05"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul05.bmp")
        Set imgModule.Picture = ImageList2.ListImages(5).Picture
        fraColumn.Enabled = False
        lstColumn.Visible = False
        cmbColumn.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        lblColShape.Visible = False
        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-05"
    Case "Module06"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul06.bmp")
        Set imgModule.Picture = ImageList2.ListImages(6).Picture
        fraColumn.Enabled = False
        lstColumn.Visible = False
        cmbColumn.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        lblColShape.Visible = False
        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-06"
    Case "Module07"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul07.bmp")
        Set imgModule.Picture = ImageList2.ListImages(7).Picture
        fraColumn.Enabled = False
        lstColumn.Visible = False
        cmbColumn.Visible = False
        fraBeam.Enabled = False
        lstBeam.Visible = False
        cmbBeam.Visible = False
        lblSpace.Visible = True
        lblSpace.Caption = "Space(m):"
        txtSpace.Visible = True
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        lblColShape.Visible = False
        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-07"
    Case "Module08"
        'imgModule.Picture = LoadPicture(App.Path & "\BMP\VB\Module\XZ_Modul08.bmp")
        Set imgModule.Picture = ImageList2.ListImages(8).Picture
        fraColumn.Enabled = False
        lstColumn.Visible = False
        cmbColumn.Visible = False
        lblSpace.Visible = False
        txtSpace.Visible = False
        cmbType.Clear
        cmbType.Text = "Type-01"
        cmbType.AddItem "Type-01"
        cmbType.AddItem "Type-02"
        cmbType.AddItem "Type-03"
        lblColShape.Visible = False
        cmbColShape.Visible = False
        fraModule.Caption = "View of Module-08"
End Select


Call Type_BMP_Control_Start


End Sub

Private Sub Type_BMP_Control_Start()

If gstr_VBdir_Flag = "Y" Then
    Set imgAxis.Picture = ImageList1.ListImages(2).Picture

Else
    Set imgAxis.Picture = ImageList1.ListImages(1).Picture

End If



Select Case gstr_VBGP_Flag
    Case "Module01"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
            Set imgModel.Picture = ImageM1.ListImages(1).Picture
        Else
            Set imgModel.Picture = ImageM1.ListImages(2).Picture
        End If
    Case "Module02"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_S.bmp")
            Set imgModel.Picture = ImageM2.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_W.bmp")
            Set imgModel.Picture = ImageM2.ListImages(2).Picture
        End If
    Case "Module03"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_W.bmp")
            Set imgModel.Picture = ImageM3.ListImages(2).Picture
        End If
        
    Case "Module04"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_W.bmp")
            Set imgModel.Picture = ImageM4.ListImages(2).Picture
        End If
        
    Case "Module05"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type01.bmp")
            Set imgModel.Picture = ImageM5.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type02.bmp")
            Set imgModel.Picture = ImageM5.ListImages(2).Picture
        End If
       
    Case "Module06"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type01.bmp")
            Set imgModel.Picture = ImageM6.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type02.bmp")
            Set imgModel.Picture = ImageM6.ListImages(2).Picture
        End If
        
    Case "Module07"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type01.bmp")
            Set imgModel.Picture = ImageM7.ListImages(1).Picture
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type02.bmp")
            Set imgModel.Picture = ImageM7.ListImages(2).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type03.bmp")
            Set imgModel.Picture = ImageM7.ListImages(3).Picture
        End If
        
    Case "Module08"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type01.bmp")
            Set imgModel.Picture = ImageM8.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type02.bmp")
            Set imgModel.Picture = ImageM8.ListImages(2).Picture
        End If
        
End Select

End Sub
Private Sub Type_BMP_Control()


Select Case gstr_VBGP_Flag
    Case "Module01"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type01_S.bmp")
            Set imgModel.Picture = ImageM1.ListImages(1).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type02_S.bmp")
            Set imgModel.Picture = ImageM1.ListImages(3).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type03_S.bmp")
            Set imgModel.Picture = ImageM1.ListImages(5).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type04_S.bmp")
            Set imgModel.Picture = ImageM1.ListImages(7).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type01_W.bmp")
            Set imgModel.Picture = ImageM1.ListImages(2).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type02_W.bmp")
            Set imgModel.Picture = ImageM1.ListImages(4).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul01_Type03_W.bmp")
            Set imgModel.Picture = ImageM1.ListImages(6).Picture
        End If
    Case "Module02"
        
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_S.bmp")
            Set imgModel.Picture = ImageM2.ListImages(1).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type02_S.bmp")
            Set imgModel.Picture = ImageM2.ListImages(3).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type03_S.bmp")
            Set imgModel.Picture = ImageM2.ListImages(5).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type04_S.bmp")
            Set imgModel.Picture = ImageM2.ListImages(7).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type01_W.bmp")
            Set imgModel.Picture = ImageM2.ListImages(2).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type02_W.bmp")
            Set imgModel.Picture = ImageM2.ListImages(4).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul02_Type03_W.bmp")
            Set imgModel.Picture = ImageM2.ListImages(6).Picture
        End If
    Case "Module03"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(1).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type02_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(3).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type03_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(5).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type04_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(7).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-05" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type05_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(8).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-06" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type06_S.bmp")
            Set imgModel.Picture = ImageM3.ListImages(9).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-07" Then
            Set imgModel.Picture = ImageM3.ListImages(10).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type01_W.bmp")
            Set imgModel.Picture = ImageM3.ListImages(2).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type02_W.bmp")
            Set imgModel.Picture = ImageM3.ListImages(4).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul03_Type03_W.bmp")
            Set imgModel.Picture = ImageM3.ListImages(6).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            Set imgModel.Picture = ImageM3.ListImages(11).Picture
        End If
        
    Case "Module04"
        If CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(1).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type02_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(3).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type03_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(5).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type04_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(7).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-05" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type05_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(8).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-06" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type06_S.bmp")
            Set imgModel.Picture = ImageM4.ListImages(9).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Strong Axis" And CStr(Trim(cmbType.Text)) = "Type-07" Then
            Set imgModel.Picture = ImageM4.ListImages(10).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type01_W.bmp")
            Set imgModel.Picture = ImageM4.ListImages(2).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type02_W.bmp")
            Set imgModel.Picture = ImageM4.ListImages(4).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-03" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul04_Type03_W.bmp")
            Set imgModel.Picture = ImageM4.ListImages(6).Picture
        ElseIf CStr(Trim(cmbColShape.Text)) = "Weak Axis" And CStr(Trim(cmbType.Text)) = "Type-04" Then
            Set imgModel.Picture = ImageM4.ListImages(11).Picture
        End If
        
    Case "Module05"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type01.bmp")
            Set imgModel.Picture = ImageM5.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul05_Type02.bmp")
            Set imgModel.Picture = ImageM5.ListImages(2).Picture
        End If
       
    Case "Module06"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type01.bmp")
            Set imgModel.Picture = ImageM6.ListImages(1).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul06_Type02.bmp")
            Set imgModel.Picture = ImageM6.ListImages(2).Picture
        End If
        
    Case "Module07"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type01.bmp")
            Set imgModel.Picture = ImageM7.ListImages(1).Picture
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type02.bmp")
            Set imgModel.Picture = ImageM7.ListImages(2).Picture
        Else
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul07_Type03.bmp")
            Set imgModel.Picture = ImageM7.ListImages(3).Picture
        End If
        
    Case "Module08"
        If CStr(Trim(cmbType.Text)) = "Type-01" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type01.bmp")
            Set imgModel.Picture = ImageM8.ListImages(1).Picture
        ElseIf CStr(Trim(cmbType.Text)) = "Type-02" Then
            'imgModel.Picture = LoadPicture(App.Path & "\BMP\VB\Type\Modul08_Type02.bmp")
            Set imgModel.Picture = ImageM8.ListImages(2).Picture
        ElseIf CStr(Trim(cmbType.Text)) = "Type-03" Then
            Set imgModel.Picture = ImageM8.ListImages(3).Picture
        End If
        
End Select

End Sub

Private Sub optFeet_Click()
lblSpace.Caption = "Space(ft):"

End Sub

Private Sub optInch_Click()
lblSpace.Caption = "Space(in):"

End Sub

Private Sub optM_Click()
lblSpace.Caption = "Space(m):"
End Sub

Private Sub optMM_Click()
lblSpace.Caption = "Space(mm):"
End Sub

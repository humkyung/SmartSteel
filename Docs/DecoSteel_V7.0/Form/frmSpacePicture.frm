VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpacePicture 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Picture"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   Icon            =   "frmSpacePicture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1620
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   235
      ImageHeight     =   259
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":2CCAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   235
      ImageHeight     =   259
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":59948
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":865E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":B3284
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":DFF22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":10CBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":13985E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":1664FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":19319A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":1BFE38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   235
      ImageHeight     =   259
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":1ECAD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":21979E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":24643C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpacePicture.frx":273104
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgPicture 
      Height          =   3885
      Left            =   0
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "frmSpacePicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If gin_Chk_Flag01 = 0 Then
            i = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
Else
            i = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End If

Select Case gs_SP_Flag
    Case "1"

        Set imgPicture.Picture = ImageList1.ListImages(1).Picture

        Me.Caption = "Picture of Space01"
    Case "2"

        Set imgPicture.Picture = ImageList1.ListImages(2).Picture
        Me.Caption = "Picture of Space02"
    Case "3"

        Set imgPicture.Picture = ImageList1.ListImages(3).Picture
        Me.Caption = "Picture of Space01"
    Case "4"

        Set imgPicture.Picture = ImageList1.ListImages(4).Picture
        Me.Caption = "Picture of Space02"
    Case "A1"

        Set imgPicture.Picture = ImageList2.ListImages(1).Picture
        Me.Caption = "Picture of Type A1"
    Case "A2"

        Set imgPicture.Picture = ImageList2.ListImages(2).Picture
        Me.Caption = "Picture of Type A2"
    Case "A3"

        Set imgPicture.Picture = ImageList2.ListImages(3).Picture
        Me.Caption = "Picture of Type A3"
    Case "A4"

        Set imgPicture.Picture = ImageList2.ListImages(4).Picture
        Me.Caption = "Picture of Type A4"
    Case "B1"
        
        Set imgPicture.Picture = ImageList2.ListImages(5).Picture
        Me.Caption = "Picture of Type B1"
    Case "B2"
        
        Set imgPicture.Picture = ImageList2.ListImages(6).Picture
        Me.Caption = "Picture of Type B2"
    Case "B3"
        
        Set imgPicture.Picture = ImageList2.ListImages(7).Picture
        Me.Caption = "Picture of Type B3"
    Case "B4"
        
        Set imgPicture.Picture = ImageList2.ListImages(8).Picture
        Me.Caption = "Picture of Type B4"
    Case "B5"
        
        Set imgPicture.Picture = ImageList2.ListImages(9).Picture
        Me.Caption = "Picture of Type B5"
    Case "I"
        
        Set imgPicture.Picture = ImageList3.ListImages(1).Picture
        Me.Caption = "Picture of Type I"
    Case "II"
        
        Set imgPicture.Picture = ImageList3.ListImages(2).Picture
        Me.Caption = "Picture of Type II"
End Select

End Sub

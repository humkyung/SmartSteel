Attribute VB_Name = "Public_Variable"
Public Const gs_FWP = "C:\win32app\INGR\FWPLUS"
Public Const gs_LGFWP = "C:\LGFWP_Certification"

Public gs_CertiKey As String
'Public gs_PMLunit As String
Public gs_PW As String
Public gs_Caption As String
'Public gstr_Job As String
Public Gdo_Cheight As Double
Public Gdo_Cwidth As Double
Public Gdo_Ctw As Double
Public Gdo_Ctf As Double
Public Gdo_CStartX As Double
Public Gdo_CStartY As Double
Public Gdo_Bheight As Double
Public Gdo_Bwidth As Double
Public Gdo_Btw As Double
Public Gdo_Btf As Double
Public Gdo_BRheight As Double
Public Gdo_BRWidth As Double
Public Gdo_BRtw As Double
Public Gdo_BRtf As Double
Public Gstr_BRtype As String
Public Gstr_Brdir As String
Public Gstr_GPtype As String
Public Gdo_GP_Width As Double
Public Gdo_GP_Depth As Double
Public Gdo_GP_Thick As Double
Public gin_Int As Integer
'BasePlate변수
Public gsin_Xlen As Single
Public gsin_Ylen As Single
Public gsin_Pthk  As Single
Public gsin_StiffThk As Single
Public gsin_Xbtob As Single
Public gsin_Xcls As Single
Public gsin_Ybtob As Single
Public gsin_Ycls As Single
Public gsin_DimF As Single
Public gsin_DimG As Single
Public gin_BoltEA As Integer
Public gstr_BoltType As String
Public gstr_BoltName As String
Public gstr_Type As String
Public gstr_Unit As String
Public gsin_BoltDia As Single
Public gsin_NutDia As Single
Public gsin_NutHei As Single
Public gstr_NutUnit As String

Public gstr_BPF_Flag As String

Public gsin_Cdepth As Single
Public gsin_Cwidth As Single
Public gsin_Cwt As Single
Public gsin_Cft As Single

Public gsin_Bedepth As Single
Public gsin_Bewidth As Single
Public gsin_Bewt As Single
Public gsin_Beft As Single

Public gsin_SubBedepth As Single
Public gsin_SubBewidth As Single
Public gsin_SubBewt As Single
Public gsin_SubBeft As Single

Public gsin_rBedepth As Single
Public gsin_rBewidth As Single
Public gsin_rBewt As Single
Public gsin_rBeft As Single

Public gsin_Bdepth As Single
Public gsin_Bwidth As Single
Public gsin_Bwt As Single
Public gsin_Bft As Single

'Vertical Gusset Plate 변수
Public gstr_VBGP_Flag As String
Public gstr_VBdir_Flag As String
Public gstr_HBGP_Flag As String
Public gstr_MCEP_Flag As String
Public gstr_MCEP_type As String
Public gstr_MCdir_Flag As String
Public gstr_SCP_Flag As String
Public gstr_SCP_type As String
Public gstr_SCdir_Flag As String

Public gsin_SP1 As Single
Public gsin_SP2 As Single
Public gsin_SP3 As Single
Public gsin_HTB_Num As Integer
Public gsin_HTB_SNum As Integer
Public gsin_HTB_Space As Single
Public gsin_GPThk As Single
Public gsin_Gage As Single

Public gs_SP_Flag As String

'Moment connection end plate 변수
Public gsin_L As Single
Public gsin_L2 As Single
Public gsin_W As Single
Public gsin_A As Single
Public gsin_B As Single
Public gsin_C As Single
Public gsin_D As Single
Public gsin_E As Single
Public gsin_F As Single
Public gsin_G As Single
Public gsin_H As Single
Public gsin_I As Single
Public gsin_J As Single
Public gsin_SPATop As Single
Public gsin_SPABot As Single

Public gsin_rPthk  As Single
Public gsin_rStiffThk As Single
Public gstr_rType As String
Public gstr_rUnit As String
Public gstr_rBoltName As String
Public gsin_rBoltDia As Single
Public gsin_rNutDia As Single
Public gsin_rNutHei As Single
Public gstr_rNutUnit As String

Public gsin_rL As Single
Public gsin_rL2 As Single
Public gsin_rW As Single
Public gsin_rA As Single
Public gsin_rB As Single
Public gsin_rC As Single
Public gsin_rD As Single
Public gsin_rE As Single
Public gsin_rF As Single
Public gsin_rG As Single
Public gsin_rH As Single
Public gsin_rI As Single
Public gsin_rJ As Single
Public gsin_rSPATop As Single
Public gsin_rSPABot As Single

Public gsin_Gap As Single

Public gstr_Shape As String
Public gin_Shape_Flag As Integer
' General
Public gstr_Grade As String
Public gstr_Material As String
Public gstr_BPClass As String
Public gstr_VBClass As String
Public gstr_HBClass As String
Public gstr_MCClass As String
Public gstr_SCClass As String

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public gin_Chk_Flag01 As Integer
Public gin_Chk_Flag02 As Integer
Public gin_MenuFlag As Integer

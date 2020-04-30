Attribute VB_Name = "Fuction"
'Public Sub vsd_int(vsdName As VSDraw) ', ByRef Ewidth As Double, ByRef Edepth As Double)
'
'With vsdName
'
'        ' start fresh
'        .Clear
'
'        ' set brush
'        .BrushStyle = bsTransparent
''        .BrushColor = RGB(210 + Rnd * 40, 210 + Rnd * 40, 210 + Rnd * 40)
'        .PenColor = RGB(0, 255, 255)
'        .PenStyle = psSolid
'
'        .DrawRectangle 0, 0, 996, 998
'
'        .PenColor = RGB(255, 255, 255)
''        Ewidth = 500
''        Edepth = 500
'
''        .DrawLine 100, 200, 100, 400
'        ' set scale
''        v = Val(txtScaleWidth)
''        If Abs(v) > 100 Then .ScaleWidth = v
''        txtScaleWidth = .ScaleWidth
''        v = Val(txtScaleHeight)
''        If Abs(v) > 100 Then .ScaleHeight = v
''        txtScaleHeight = .ScaleHeight
'
'        ' set font
''        .FontBold = False
''        .FontItalic = False
''        .FontName = cmbFontName.List(cmbFontName.ListIndex)
''        v = Val(cmbFontSize.List(cmbFontSize.ListIndex))
''        .FontSize = v * .ScaleHeight / 100
''        .KeepTextAspect = cmbKeep.Value
'
'        ' draw a box to show text dimensions
''        Dim txtWid!, txtHei!
''        txtWid = .TextWidth(txtText)
''        txtHei = .TextHeight(txtText)
''        .X1 = .ScaleWidth / 2
''        .Y1 = .ScaleHeight / 2 - .FontSize / 2
''        .DrawRectangle .X1 - txtWid / 2, .Y1, .X1 + txtWid / 2, .Y1 + .FontSize
'
'        ' draw the text over the box
''        .TextAlign = tadCenter
''        .Text = txtText
'
'        ' draw a few more boxes around the corners to show the scale
''        .DrawRectangle 0, 0, 100, 100
''        .DrawRectangle 0, .ScaleHeight - 100, 100, .ScaleHeight
''        .DrawRectangle .ScaleWidth - 100, 0, .ScaleWidth, 100
''        .DrawRectangle .ScaleWidth - 100, .ScaleHeight - 100, .ScaleWidth, .ScaleHeight
'
'    End With
'End Sub

'Public Sub vsd_column(vsdName As VSDraw, H As Double, W As Double, Tw As Double, Tf As Double, _
'                      OPx As Double, OPy As Double, Clength As Double)
'
'With vsdName
'    .DrawRectangle OPx, OPy, OPx + H, OPy + Clength
'    .DrawLine OPx + Tf, OPy, OPx + Tf, OPy + Clength
'    .DrawLine OPx + H - Tf, OPy, OPx + H - Tf, OPy + Clength
'    .PenStyle = psDashDot
'    .PenColor = RGB(0, 255, 0)
'    .DrawLine OPx + H / 2, OPy - 50, OPx + H / 2, OPy + Clength + 50
'    .PenColor = RGB(255, 255, 255)
'    .PenStyle = psSolid
'
'End With
'
'End Sub
'Public Sub vsd_Beam(vsdName As VSDraw, H As Double, W As Double, Tw As Double, Tf As Double, _
'                      OPx As Double, OPy As Double, Blength As Double, Ch As Double)
'
'With vsdName
'    .DrawRectangle OPx, OPy, OPx + Blength, OPy + H
'    .DrawLine OPx, OPy + Tf, OPx + Blength, OPy + Tf
'    .DrawLine OPx, OPy + H - Tf, OPx + Blength, OPy + H - Tf
'    .PenStyle = psDashDot
'    .PenColor = RGB(0, 255, 0)
'    .DrawLine OPx - Ch - 50, OPy + H / 2, OPx + Blength + 50, OPy + H / 2
'    .PenColor = RGB(255, 255, 255)
'    .PenStyle = psSolid
'
'End With
'
'End Sub
'
'Public Sub vsd_Bracing(vsdName As VSDraw, Ch As Double, Bh As Double, Clength As Double, _
'                       OPx As Double, OPy As Double, Mtype As String, BRh As Double, BRw As Double, _
'                       Brtf As Double, Brtw As Double, Optional SFlag As String)
'Dim tX1 As Double, tY1 As Double, tX2 As Double, tY2 As Double
'Dim tX3 As Double, tY3 As Double, tX4 As Double, tY4 As Double
'Dim bX1 As Double, bY1 As Double, bX2 As Double, bY2 As Double
'Dim bX3 As Double, bY3 As Double, bX4 As Double, bY4 As Double
'Dim TanAngle As Double
'With vsdName
'    Select Case Mtype
'        Case "Channel"
'            tX1 = OPx + Ch + Sqr(BRh ^ 2 / 2)
'            tY1 = OPy + Bh
'            tX2 = tX1 + 350
'            tY2 = tY1 + 350
'
'            bX1 = OPx + Ch
'            bY1 = OPy + Bh + Sqr(BRh ^ 2 / 2)
'            bX2 = bX1 + 350
'            bY2 = bY1 + 350
'            .DrawLine tX1, tY1, tX2, tY2
'            .DrawLine bX1, bY1, bX2, bY2
'            .DrawLine tX1, tY1, bX1, bY1
'            .DrawLine tX2, tY2, bX2, bY2
'
'            If SFlag = "Dir1" Then
'                tX3 = tX1 - Sqr(Brtf ^ 2 / 2) * 2
'                tY3 = tY1 + Sqr(Brtf ^ 2 / 2) * 2
'                tX4 = tX3 + 350
'                tY4 = tY3 + 350
'                .DrawLine tX3, tY3, tX4, tY4
'
'                bX3 = bX1 + Sqr(Brtf ^ 2 / 2) * 2
'                bY3 = bY1 - Sqr(Brtf ^ 2 / 2) * 2
'                bX4 = bX3 + 350
'                bY4 = bY3 + 350
'                .DrawLine bX3, bY3, bX4, bY4
'
'            ElseIf SFlag = "Dir2" Then
'                .PenStyle = psDot
'                tX3 = tX1 - Sqr(Brtf ^ 2 / 2) * 2
'                tY3 = tY1 + Sqr(Brtf ^ 2 / 2) * 2
'                tX4 = tX3 + 350
'                tY4 = tY3 + 350
'                .DrawLine tX3, tY3, tX4, tY4
'
'                bX3 = bX1 + Sqr(Brtf ^ 2 / 2) * 2
'                bY3 = bY1 - Sqr(Brtf ^ 2 / 2) * 2
'                bX4 = bX3 + 350
'                bY4 = bY3 + 350
'                .DrawLine bX3, bY3, bX4, bY4
'                .PenStyle = psSolid
'            End If
'        Case "Angle"
'            tX1 = OPx + Ch + Sqr(BRh ^ 2 / 2)
'            tY1 = OPy + Bh
'            tX2 = tX1 + 350
'            tY2 = tY1 + 350
'
'            bX1 = OPx + Ch
'            bY1 = OPy + Bh + Sqr(BRh ^ 2 / 2)
'            bX2 = bX1 + 350
'            bY2 = bY1 + 350
'            .DrawLine tX1, tY1, tX2, tY2
'            .DrawLine bX1, bY1, bX2, bY2
'            .DrawLine tX1, tY1, bX1, bY1
'            .DrawLine tX2, tY2, bX2, bY2
'
'            If SFlag = "Dir1" Then
'                tX3 = tX1 - Sqr(Brtf ^ 2 / 2) * 2
'                tY3 = tY1 + Sqr(Brtf ^ 2 / 2) * 2
'                tX4 = tX3 + 350
'                tY4 = tY3 + 350
'                .DrawLine tX3, tY3, tX4, tY4
'
'            ElseIf SFlag = "Dir2" Then
'                .PenStyle = psDot
'                tX3 = tX1 - Sqr(Brtf ^ 2 / 2) * 2
'                tY3 = tY1 + Sqr(Brtf ^ 2 / 2) * 2
'                tX4 = tX3 + 350
'                tY4 = tY3 + 350
'                .DrawLine tX3, tY3, tX4, tY4
'
'                .PenStyle = psSolid
'            End If
'        Case "Double Angle"
'
'        Case "Double Cannel"
'
'        Case "Tee"
'
'    End Select
''    TanAngle = Atn(BH / CH) * (180 / 3.14159265358979)
''
''    .PenStyle = psDashDot
''    .DrawLine OPx + CH / 2, OPy + BH / 2, OPx + Clength, OPy + Clength * Tan((TanAngle * 3.14159265358979) / 180)
''    .PenStyle = psSolid
'
'End With
'
'End Sub
'
'Public Sub vsd_Draw_View_01(vsdName As VSDraw, SF As Double, Ch As Double, Cw As Double, _
'                            Ctw As Double, Ctf As Double, _
'                            Cl As Double, Cx As Double, Cy As Double, Bh As Double, _
'                            Bw As Double, Btw As Double, Btf As Double, Bl As Double, _
'                            BRh As Double, BRw As Double, _
'                            Brtw As Double, Brtf As Double, BrType As String, _
'                            Optional Flag As String)
'
'Dim lv_Column_h As Double, lv_Column_w As Double, lv_Column_tw As Double, lv_Column_tf As Double
'Dim lv_Column_Length As Double, lv_Column_X As Double, lv_Column_Y As Double
'Dim lv_Beam_h As Double, lv_Beam_w As Double, lv_Beam_tw As Double, lv_Beam_tf As Double
'Dim lv_Beam_Length As Double, lv_Beam_X As Double, lv_Beam_Y As Double
'Dim lv_Bracing_h As Double, lv_Bracing_w As Double, lv_Bracing_tw As Double, lv_Bracing_tf As Double
'
'vsdName.Clear
'
'lv_Column_h = Ch * SF
'lv_Column_w = Cw * SF
'lv_Column_tw = Ctw * (SF + 0.2)
'lv_Column_tf = Ctf * (SF + 0.2)
'lv_Column_Length = Cl
'lv_Column_X = Cx: lv_Column_Y = Cy
'
'lv_Beam_h = Bh * SF
'lv_Beam_w = Bw * ScaleF
'lv_Beam_tw = Btw * (SF + 0.2)
'lv_Beam_tf = Btf * (SF + 0.2)
'lv_Beam_Length = Bl
'lv_Beam_X = lv_Column_X + lv_Column_h
'lv_Beam_Y = lv_Column_Y
'
'lv_Bracing_h = BRh * SF
'lv_Bracing_w = BRw * SF
'lv_Bracing_tf = Brtf * SF
'lv_braicng_tw = Brtw * SF
'
'
'Call vsd_column(vsdName, lv_Column_h, lv_Column_w, lv_Column_tw, _
'                     lv_Column_tf, lv_Column_X, lv_Column_Y, lv_Column_Length)
'Call vsd_Beam(vsdName, lv_Beam_h, lv_Beam_w, lv_Beam_tw, _
'                   lv_Beam_tf, lv_Beam_X, lv_Beam_Y, lv_Beam_Length, lv_Column_h)
'
'Call vsd_Bracing(vsdName, lv_Column_h, lv_Beam_h, lv_Column_Length, _
'                      lv_Column_X, lv_Column_Y, BrType, _
'                      lv_Bracing_h, lv_Bracing_w, lv_Bracing_tf, lv_Bracing_tw, Flag)
'
'
'
'
'End Sub
'Public Sub vsd_Draw_GassetPlate_01(vsdName As VSDraw, SF As Double, OriginX As Double, _
'                                   OriginY As Double, Ch As Double, Bh As Double, _
'                                   GP_Width As Double, GP_Depth As Double, _
'                                   BR_Type As String, BR_Dir As String)
'
'Dim lv_Bema_X As Double, lv_Beam_Y As Double
'Dim GP_X1 As Double, GP_Y1 As Double, GP_X2 As Double, GP_Y2 As Double
'
'lv_Beam_X = OriginX + Ch * SF
'lv_Beam_Y = OriginY + Bh * SF
'
'GP_X1 = lv_Beam_X
'GP_Y1 = lv_Beam_Y
'GP_X2 = GP_X1 + SF * GP_Width
'GP_Y2 = GP_Y1 + SF * GP_Depth
'
'
'With vsdName
'    .PenColor = RGB(250, 0, 250)
'    Select Case BR_Type
'        Case "Channel"
'            If BR_Dir = "Dir1" Then
'                .PenStyle = psDot
'                .DrawRectangle GP_X1, GP_Y1, GP_X2, GP_Y2
'                .PenStyle = psSolid
'            Else
'                .BrushStyle = bsDiagonalDown
'                .BrushColor = RGB(250, 0, 250)
'                .DrawRectangle GP_X1, GP_Y1, GP_X2, GP_Y2
'                .BrushColor = RGB(250, 250, 250)
'                .BrushStyle = bsTransparent
'            End If
'        Case "Angle"
'            If BR_Dir = "Dir1" Then
'                .PenStyle = psDot
'                .DrawRectangle GP_X1, GP_Y1, GP_X2, GP_Y2
'                .PenStyle = psSolid
'            Else
'                .BrushStyle = bsDiagonalDown
'                .BrushColor = RGB(250, 0, 250)
'                .DrawRectangle GP_X1, GP_Y1, GP_X2, GP_Y2
'                .BrushColor = RGB(250, 250, 250)
'                .BrushStyle = bsTransparent
'            End If
'
'    End Select
'    .PenColor = RGB(250, 250, 250)
'
'End With
'
'End Sub
Public Sub PML_Print()
Dim TempPath As String

frmModule1.ComDialog.DialogTitle = "Save PML File "
frmModule1.ComDialog.Filter = "TEST (*.pml)|*.pml|"
frmModule1.ComDialog.FileName = "Test.pml"
frmModule1.ComDialog.ShowSave
 
TempPath = frmModule1.ComDialog.FileName

' Gstr_BRtype
' Gstr_Brdir
' Gstr_GPtype
Select Case Gstr_GPtype
    Case "Type01"
        Call GassetPlate_PML_Type01(TempPath, Gdo_Cheight, Gdo_Bheight)
    Case "Type02"
        Call GassetPlate_PML_Type02(TempPath, Gdo_Cheight, Gdo_Bheight)
    Case Else
        MsgBox "Data Input Error...."
        Exit Sub
End Select




End Sub
Public Sub GassetPlate_PML_Type01(subPath As String, subCH As Double, subBH As Double)

Dim State_01 As String
Dim State_02 As String
Dim State_03 As String
Dim State_04 As String
Dim Ox As Double, Oy As Double, Oz As Double
Dim X As Double, Y As Double, Z As Double


State_01 = "origin prompt = " & """Pick Start Point""" & ";"
State_02 = "var_type = " & """float""" & ";"
State_03 = "material = " & """concrete"""
State_04 = "grade = " & """FC_4"""
'Gdo_GP_Width
'Gdo_GP_Depth
'Gdo_GP_Thick
Ox = subCH
Oy = 0
Oz = -subBH

X = Ox + Gdo_GP_Width
Y = Oy - Gdo_GP_Thick
Z = Oz - Gdo_GP_Depth


Open subPath For Output As #1
    Print #1, State_01
    Print #1, "assign endx = %%point_x, " & State_02
    Print #1, "assign endy = %%point_y, " & State_02
    Print #1, "assign endz = %%point_z, " & State_02
    Print #1, "origin local = endx, endy, endz;"
    Print #1, "plc_volume " & State_03 & ", " & State_04 & ","
    Print #1, "class = 0, fab_type = 0,"
    Print #1, "vert1 = " & Ox & ", " & Oy & ", " & Oz & ","
    Print #1, "vert2 = " & X & ", " & Oy & ", " & Oz & ","
    Print #1, "vert3 = " & X & ", " & Oy & ", " & Z & ","
    Print #1, "vert4 = " & Ox & ", " & Oy & ", " & Z & ","
    Print #1, "vert5 = " & Ox & ", " & Y & ", " & Oz & ","
    Print #1, "vert6 = " & X & ", " & Y & ", " & Oz & ","
    Print #1, "vert7 = " & X & ", " & Y & ", " & Z & ","
    Print #1, "vert8 = " & Ox & ", " & Y & ", " & Z & ";"
Close #1


End Sub
Public Sub GassetPlate_PML_Type02(subPath As String, subCH As Double, subBH As Double)

Dim State_01 As String
Dim State_02 As String
Dim State_03 As String
Dim State_04 As String
Dim Ox As Double, Oy As Double, Oz As Double
Dim X As Double, Y As Double, Z As Double


State_01 = "origin prompt = " & """Pick Start Point""" & ";"
State_02 = "var_type = " & """float""" & ";"
State_03 = "material = " & """concrete"""
State_04 = "grade = " & """FC_4"""
Ox = subCH
Oy = 0
Oz = -subBH

X = Ox + Gdo_GP_Width
Y = Oy + Gdo_GP_Thick
Z = Oz - Gdo_GP_Depth


Open subPath For Output As #1
    Print #1, State_01
    Print #1, "assign endx = %%point_x, " & State_02
    Print #1, "assign endy = %%point_y, " & State_02
    Print #1, "assign endz = %%point_z, " & State_02
    Print #1, "origin local = endx, endy, endz;"
    Print #1, "plc_volume " & State_03 & ", " & State_04 & ","
    Print #1, "class = 0, fab_type = 0,"
    Print #1, "vert1 = " & Ox & ", " & Oy & ", " & Oz & ","
    Print #1, "vert2 = " & X & ", " & Oy & ", " & Oz & ","
    Print #1, "vert3 = " & X & ", " & Oy & ", " & Z & ","
    Print #1, "vert4 = " & Ox & ", " & Oy & ", " & Z & ","
    Print #1, "vert5 = " & Ox & ", " & Y & ", " & Oz & ","
    Print #1, "vert6 = " & X & ", " & Y & ", " & Oz & ","
    Print #1, "vert7 = " & X & ", " & Y & ", " & Z & ","
    Print #1, "vert8 = " & Ox & ", " & Y & ", " & Z & ";"
Close #1


End Sub

Public Function Function_Scale(Flag1 As Double, Flag2 As Double) As Double


If Flag1 >= Flag2 Then
    If Flag1 < 100 Then
        Function_Scale = 1
    ElseIf Flag1 >= 100 And Flag1 <= 200 Then
        Function_Scale = 0.5
    ElseIf Flag1 >= 200 And Flag1 <= 300 Then
        Function_Scale = 0.4
    ElseIf Flag1 >= 300 And Flag1 <= 400 Then
        Function_Scale = 0.3
    ElseIf Flag1 > 400 And Flag1 <= 500 Then
        Function_Scale = 0.2
    ElseIf Flag1 > 500 And Flag1 <= 600 Then
        Function_Scale = 0.1
    ElseIf Flag1 > 600 And Flag1 <= 700 Then
        Function_Scale = 0.1
    ElseIf Flag1 > 700 And Flag1 <= 800 Then
        Function_Scale = 0.1
    ElseIf Flag1 > 800 And Flag1 <= 900 Then
        Function_Scale = 0.1
    ElseIf Flag1 > 900 Then
        Function_Scale = 0.1
    End If
Else
    If Flag2 < 100 Then
        Function_Scale = 1
    ElseIf Flag2 >= 100 And Flag2 <= 200 Then
        Function_Scale = 0.5
    ElseIf Flag2 >= 200 And Flag2 <= 300 Then
        Function_Scale = 0.4
    ElseIf Flag2 >= 300 And Flag2 <= 400 Then
        Function_Scale = 0.3
    ElseIf Flag2 > 400 And Flag2 <= 500 Then
        Function_Scale = 0.2
    ElseIf Flag2 > 500 And Flag2 <= 600 Then
        Function_Scale = 0.1
    ElseIf Flag2 > 600 And Flag2 <= 700 Then
        Function_Scale = 0.1
    ElseIf Flag2 > 700 And Flag2 <= 800 Then
        Function_Scale = 0.1
    ElseIf Flag2 > 800 And Flag2 <= 900 Then
        Function_Scale = 0.1
    ElseIf Flag2 > 900 Then
        Function_Scale = 0.1
    End If
End If
End Function

Public Sub BP_Fixed_PML_Type01_X(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single, subF As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------"

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    Print #1, "assign DimF = " & CStr(subF) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------"

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal
    Print #1, "assign WRY11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY12 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY13 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY14 =  WRY11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRY21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY22 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY23 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY24 =  -WRY11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal
    Print #1, "assign SRX11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX12 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX13 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX14 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY1 =  DimF , var_type=""Float"";"

    Print #1, "assign SRX21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX22 =  -SRX11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX23 =  -SRX11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX24 =  -SRX11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY2 =  DimF , var_type=""Float"";"

    ' !BASE PLATE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = WRX1-RBt, WRY11, WRZ11,"
    Print #1, "vert2 = WRX1-RBt, WRY12, WRZ12,"
    Print #1, "vert3 = WRX1-RBt, WRY13, WRZ13,"
    Print #1, "vert4 = WRX1-RBt, WRY14, WRZ14,"
    Print #1, "vert5 = WRX1-RBt, WRY15, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX2, WRY21, WRZ21,"
    Print #1, "vert2 = WRX2, WRY22, WRZ22,"
    Print #1, "vert3 = WRX2, WRY23, WRZ23,"
    Print #1, "vert4 = WRX2, WRY24, WRZ24,"
    Print #1, "vert5 = WRX2, WRY25, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX2+RBt, WRY21, WRZ21,"
    Print #1, "vert2 = -WRX2+RBt, WRY22, WRZ22,"
    Print #1, "vert3 = -WRX2+RBt, WRY23, WRZ23,"
    Print #1, "vert4 = -WRX2+RBt, WRY24, WRZ24,"
    Print #1, "vert5 = -WRX2+RBt, WRY25, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX1, WRY11, WRZ11,"
    Print #1, "vert2 = -WRX1, WRY12, WRZ12,"
    Print #1, "vert3 = -WRX1, WRY13, WRZ13,"
    Print #1, "vert4 = -WRX1, WRY14, WRZ14,"
    Print #1, "vert5 = -WRX1, WRY15, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' Strong Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBtH, SRZ11,"
    Print #1, "vert2 = SRX12, SRY1+RBtH, SRZ12,"
    Print #1, "vert3 = SRX13, SRY1+RBtH, SRZ13,"
    Print #1, "vert4 = SRX14, SRY1+RBtH, SRZ14,"
    Print #1, "vert5 = SRX15, SRY1+RBtH, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, SRY2-RBtH, SRZ21,"
    Print #1, "vert2 = SRX22, SRY2-RBtH, SRZ22,"
    Print #1, "vert3 = SRX23, SRY2-RBtH, SRZ23,"
    Print #1, "vert4 = SRX24, SRY2-RBtH, SRZ24,"
    Print #1, "vert5 = SRX25, SRY2-RBtH, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1+RBtH, SRZ11,"
    Print #1, "vert2 = SRX12, -SRY1+RBtH, SRZ12,"
    Print #1, "vert3 = SRX13, -SRY1+RBtH, SRZ13,"
    Print #1, "vert4 = SRX14, -SRY1+RBtH, SRZ14,"
    Print #1, "vert5 = SRX15, -SRY1+RBtH, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, -SRY2-RBtH, SRZ21,"
    Print #1, "vert2 = SRX22, -SRY2-RBtH, SRZ22,"
    Print #1, "vert3 = SRX23, -SRY2-RBtH, SRZ23,"
    Print #1, "vert4 = SRX24, -SRY2-RBtH, SRZ24,"
    Print #1, "vert5 = SRX25, -SRY2-RBtH, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type01_Y(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single, subF As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------"

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    Print #1, "assign DimF = " & CStr(subF) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------"

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal"
    Print #1, "assign WRX11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX12 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX13 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX14 =  WRX11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRX21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX22 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX23 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX24 =  -WRX11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal"
    Print #1, "assign SRY11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY12 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY13 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY14 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX1 =  DimF , var_type=""Float"";"

    Print #1, "assign SRY21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY22 =  -SRY11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY23 =  -SRY11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY24 =  -SRY11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX2 =  DimF , var_type=""Float"";"

    ' !BASE PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = WRX11, WRY1, WRZ11,"
    Print #1, "vert2 = WRX12, WRY1, WRZ12,"
    Print #1, "vert3 = WRX13, WRY1, WRZ13,"
    Print #1, "vert4 = WRX14, WRY1, WRZ14,"
    Print #1, "vert5 = WRX15, WRY1, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21, WRY2-RBt, WRZ21,"
    Print #1, "vert2 = WRX22, WRY2-RBt, WRZ22,"
    Print #1, "vert3 = WRX23, WRY2-RBt, WRZ23,"
    Print #1, "vert4 = WRX24, WRY2-RBt, WRZ24,"
    Print #1, "vert5 = WRX25, WRY2-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21, -WRY2, WRZ21,"
    Print #1, "vert2 = WRX22, -WRY2, WRZ22,"
    Print #1, "vert3 = WRX23, -WRY2, WRZ23,"
    Print #1, "vert4 = WRX24, -WRY2, WRZ24,"
    Print #1, "vert5 = WRX25, -WRY2, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX11, -WRY1+RBt, WRZ11,"
    Print #1, "vert2 = WRX12, -WRY1+RBt, WRZ12,"
    Print #1, "vert3 = WRX13, -WRY1+RBt, WRZ13,"
    Print #1, "vert4 = WRX14, -WRY1+RBt, WRZ14,"
    Print #1, "vert5 = WRX15, -WRY1+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' !Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX1-RBtH, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1-RBtH, SRY12, SRZ12,"
    Print #1, "vert3 = SRX1-RBtH, SRY13, SRZ13,"
    Print #1, "vert4 = SRX1-RBtH, SRY14, SRZ14,"
    Print #1, "vert5 = SRX1-RBtH, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX2+RBtH, SRY21, SRZ21,"
    Print #1, "vert2 = SRX2+RBtH, SRY22, SRZ22,"
    Print #1, "vert3 = SRX2+RBtH, SRY23, SRZ23,"
    Print #1, "vert4 = SRX2+RBtH, SRY24, SRZ24,"
    Print #1, "vert5 = SRX2+RBtH, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBtH, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBtH, SRY12, SRZ12,"
    Print #1, "vert3 = -SRX1-RBtH, SRY13, SRZ13,"
    Print #1, "vert4 = -SRX1-RBtH, SRY14, SRZ14,"
    Print #1, "vert5 = -SRX1-RBtH, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX2+RBtH, SRY21, SRZ21,"
    Print #1, "vert2 = -SRX2+RBtH, SRY22, SRZ22,"
    Print #1, "vert3 = -SRX2+RBtH, SRY23, SRZ23,"
    Print #1, "vert4 = -SRX2+RBtH, SRY24, SRZ24,"
    Print #1, "vert5 = -SRX2+RBtH, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type02_X(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single, subG As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------"

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    Print #1, "assign DimG = " & CStr(subG) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"
    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal
    Print #1, "assign WRY11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY12 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY13 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY14 =  WRY11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRY21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY22 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY23 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY24 =  -WRY11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal
    Print #1, "assign SRX11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX12 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX13 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX14 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRX21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX22 =  -SRX11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX23 =  -SRX11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX24 =  -SRX11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY2 =  DimG , var_type=""Float"";"

    Print #1, "assign SRX31 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ31 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX32 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ32 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX33 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ33 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX34 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ34 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX35 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ35 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY3 =  RBtH , var_type=""Float"";"

    ' !BASE PLATE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = WRX1-RBt, WRY11, WRZ11,"
    Print #1, "vert2 = WRX1-RBt, WRY12, WRZ12,"
    Print #1, "vert3 = WRX1-RBt, WRY13, WRZ13,"
    Print #1, "vert4 = WRX1-RBt, WRY14, WRZ14,"
    Print #1, "vert5 = WRX1-RBt, WRY15, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX2, WRY21, WRZ21,"
    Print #1, "vert2 = WRX2, WRY22, WRZ22,"
    Print #1, "vert3 = WRX2, WRY23, WRZ23,"
    Print #1, "vert4 = WRX2, WRY24, WRZ24,"
    Print #1, "vert5 = WRX2, WRY25, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX2+RBt, WRY21, WRZ21,"
    Print #1, "vert2 = -WRX2+RBt, WRY22, WRZ22,"
    Print #1, "vert3 = -WRX2+RBt, WRY23, WRZ23,"
    Print #1, "vert4 = -WRX2+RBt, WRY24, WRZ24,"
    Print #1, "vert5 = -WRX2+RBt, WRY25, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX1, WRY11, WRZ11,"
    Print #1, "vert2 = -WRX1, WRY12, WRZ12,"
    Print #1, "vert3 = -WRX1, WRY13, WRZ13,"
    Print #1, "vert4 = -WRX1, WRY14, WRZ14,"
    Print #1, "vert5 = -WRX1, WRY15, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "!Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBtH, SRZ11,"
    Print #1, "vert2 = SRX12, SRY1+RBtH, SRZ12,"
    Print #1, "vert3 = SRX13, SRY1+RBtH, SRZ13,"
    Print #1, "vert4 = SRX14, SRY1+RBtH, SRZ14,"
    Print #1, "vert5 = SRX15, SRY1+RBtH, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, SRY2-RBtH, SRZ21,"
    Print #1, "vert2 = SRX22, SRY2-RBtH, SRZ22,"
    Print #1, "vert3 = SRX23, SRY2-RBtH, SRZ23,"
    Print #1, "vert4 = SRX24, SRY2-RBtH, SRZ24,"
    Print #1, "vert5 = SRX25, SRY2-RBtH, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1+RBtH, SRZ11,"
    Print #1, "vert2 = SRX12, -SRY1+RBtH, SRZ12,"
    Print #1, "vert3 = SRX13, -SRY1+RBtH, SRZ13,"
    Print #1, "vert4 = SRX14, -SRY1+RBtH, SRZ14,"
    Print #1, "vert5 = SRX15, -SRY1+RBtH, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, -SRY2-RBtH, SRZ21,"
    Print #1, "vert2 = SRX22, -SRY2-RBtH, SRZ22,"
    Print #1, "vert3 = SRX23, -SRY2-RBtH, SRZ23,"
    Print #1, "vert4 = SRX24, -SRY2-RBtH, SRZ24,"
    Print #1, "vert5 = SRX25, -SRY2-RBtH, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX31, SRY3, SRZ31,"
    Print #1, "vert2 = SRX32, SRY3, SRZ32,"
    Print #1, "vert3 = SRX33, SRY3, SRZ33,"
    Print #1, "vert4 = SRX34, SRY3, SRZ34,"
    Print #1, "vert5 = SRX35, SRY3, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX31, -SRY3, SRZ31,"
    Print #1, "vert2 = -SRX32, -SRY3, SRZ32,"
    Print #1, "vert3 = -SRX33, -SRY3, SRZ33,"
    Print #1, "vert4 = -SRX34, -SRY3, SRZ34,"
    Print #1, "vert5 = -SRX35, -SRY3, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type02_Y(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single, subG As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------"

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    Print #1, "assign DimG = " & CStr(subG) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------"

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal"
    Print #1, "assign WRX11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX12 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX13 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX14 =  WRX11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRX21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX22 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX23 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX24 =  -WRX11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal"
    Print #1, "assign SRY11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY12 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY13 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY14 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRY21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY22 =  -SRY11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY23 =  -SRY11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY24 =  -SRY11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX2 =  DimG , var_type=""Float"";"

    Print #1, "assign SRY31 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ31 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY32 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ32 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY33 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ33 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY34 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ34 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY35 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ35 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX3 =  -RBtH, var_type=""Float"";"

    ' !BASE PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = WRX11, WRY1, WRZ11,"
    Print #1, "vert2 = WRX12, WRY1, WRZ12,"
    Print #1, "vert3 = WRX13, WRY1, WRZ13,"
    Print #1, "vert4 = WRX14, WRY1, WRZ14,"
    Print #1, "vert5 = WRX15, WRY1, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21, WRY2-RBt, WRZ21,"
    Print #1, "vert2 = WRX22, WRY2-RBt, WRZ22,"
    Print #1, "vert3 = WRX23, WRY2-RBt, WRZ23,"
    Print #1, "vert4 = WRX24, WRY2-RBt, WRZ24,"
    Print #1, "vert5 = WRX25, WRY2-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21, -WRY2, WRZ21,"
    Print #1, "vert2 = WRX22, -WRY2, WRZ22,"
    Print #1, "vert3 = WRX23, -WRY2, WRZ23,"
    Print #1, "vert4 = WRX24, -WRY2, WRZ24,"
    Print #1, "vert5 = WRX25, -WRY2, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX11, -WRY1+RBt, WRZ11,"
    Print #1, "vert2 = WRX12, -WRY1+RBt, WRZ12,"
    Print #1, "vert3 = WRX13, -WRY1+RBt, WRZ13,"
    Print #1, "vert4 = WRX14, -WRY1+RBt, WRZ14,"
    Print #1, "vert5 = WRX15, -WRY1+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' !Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX1-RBtH, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1-RBtH, SRY12, SRZ12,"
    Print #1, "vert3 = SRX1-RBtH, SRY13, SRZ13,"
    Print #1, "vert4 = SRX1-RBtH, SRY14, SRZ14,"
    Print #1, "vert5 = SRX1-RBtH, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX2+RBtH, SRY21, SRZ21,"
    Print #1, "vert2 = SRX2+RBtH, SRY22, SRZ22,"
    Print #1, "vert3 = SRX2+RBtH, SRY23, SRZ23,"
    Print #1, "vert4 = SRX2+RBtH, SRY24, SRZ24,"
    Print #1, "vert5 = SRX2+RBtH, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBtH, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBtH, SRY12, SRZ12,"
    Print #1, "vert3 = -SRX1-RBtH, SRY13, SRZ13,"
    Print #1, "vert4 = -SRX1-RBtH, SRY14, SRZ14,"
    Print #1, "vert5 = -SRX1-RBtH, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX2+RBtH, SRY21, SRZ21,"
    Print #1, "vert2 = -SRX2+RBtH, SRY22, SRZ22,"
    Print #1, "vert3 = -SRX2+RBtH, SRY23, SRZ23,"
    Print #1, "vert4 = -SRX2+RBtH, SRY24, SRZ24,"
    Print #1, "vert5 = -SRX2+RBtH, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX3, SRY31, SRZ31,"
    Print #1, "vert2 = SRX3, SRY32, SRZ32,"
    Print #1, "vert3 = SRX3, SRY33, SRZ33,"
    Print #1, "vert4 = SRX3, SRY34, SRZ34,"
    Print #1, "vert5 = SRX3, SRY35, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX3, -SRY31, SRZ31,"
    Print #1, "vert2 = -SRX3, -SRY32, SRZ32,"
    Print #1, "vert3 = -SRX3, -SRY33, SRZ33,"
    Print #1, "vert4 = -SRX3, -SRY34, SRZ34,"
    Print #1, "vert5 = -SRX3, -SRY35, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type03_X(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"
    Print #1, "assign DimG = HBw/2, var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal
    Print #1, "assign WRY11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY12 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY13 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY14 =  WRY11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRY21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY22 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY23 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY24 =  -WRY11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX2 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRX3 =  -RBtH , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal
    Print #1, "assign SRX11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX12 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX13 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX14 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRX21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX22 =  -SRX11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX23 =  -SRX11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX24 =  -SRX11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY2 =  DimG , var_type=""Float"";"

    Print #1, "assign SRX31 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ31 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX32 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ32 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX33 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ33 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX34 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ34 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX35 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ35 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY3 =  RBtH , var_type=""Float"";"

    ' !BASE PLATE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = WRX1-RBt, WRY11+RBt, WRZ11,"
    Print #1, "vert2 = WRX1-RBt, WRY12, WRZ12,"
    Print #1, "vert3 = WRX1-RBt, WRY13, WRZ13,"
    Print #1, "vert4 = WRX1-RBt, WRY14+RBt, WRZ14,"
    Print #1, "vert5 = WRX1-RBt, WRY15+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX2, WRY21-RBt, WRZ21,"
    Print #1, "vert2 = WRX2, WRY22, WRZ22,"
    Print #1, "vert3 = WRX2, WRY23, WRZ23,"
    Print #1, "vert4 = WRX2, WRY24-RBt, WRZ24,"
    Print #1, "vert5 = WRX2, WRY25-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX2+RBt, WRY21-RBt, WRZ21,"
    Print #1, "vert2 = -WRX2+RBt, WRY22, WRZ22,"
    Print #1, "vert3 = -WRX2+RBt, WRY23, WRZ23,"
    Print #1, "vert4 = -WRX2+RBt, WRY24-RBt, WRZ24,"
    Print #1, "vert5 = -WRX2+RBt, WRY25-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX1, WRY11+RBt, WRZ11,"
    Print #1, "vert2 = -WRX1, WRY12, WRZ12,"
    Print #1, "vert3 = -WRX1, WRY13, WRZ13,"
    Print #1, "vert4 = -WRX1, WRY14+RBt, WRZ14,"
    Print #1, "vert5 = -WRX1, WRY15+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX3, WRY11+RBt, WRZ11,"
    Print #1, "vert2 = WRX3, WRY12, WRZ12,"
    Print #1, "vert3 = WRX3, WRY13, WRZ13,"
    Print #1, "vert4 = WRX3, WRY14+RBt, WRZ14,"
    Print #1, "vert5 = WRX3, WRY15+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX3, -WRY11-RBt, WRZ11,"
    Print #1, "vert2 = -WRX3, -WRY12, WRZ12,"
    Print #1, "vert3 = -WRX3, -WRY13, WRZ13,"
    Print #1, "vert4 = -WRX3, -WRY14-RBt, WRZ14,"
    Print #1, "vert5 = -WRX3, -WRY15-RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBt, SRZ11,"
    Print #1, "vert2 = SRX12, SRY1+RBt, SRZ12,"
    Print #1, "vert3 = SRX13, SRY1+RBt, SRZ13,"
    Print #1, "vert4 = SRX14, SRY1+RBt, SRZ14,"
    Print #1, "vert5 = SRX15, SRY1+RBt, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, SRY2, SRZ21,"
    Print #1, "vert2 = SRX22, SRY2, SRZ22,"
    Print #1, "vert3 = SRX23, SRY2, SRZ23,"
    Print #1, "vert4 = SRX24, SRY2, SRZ24,"
    Print #1, "vert5 = SRX25, SRY2, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1, SRZ11,"
    Print #1, "vert2 = SRX12, -SRY1, SRZ12,"
    Print #1, "vert3 = SRX13, -SRY1, SRZ13,"
    Print #1, "vert4 = SRX14, -SRY1, SRZ14,"
    Print #1, "vert5 = SRX15, -SRY1, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, -SRY2-RBt, SRZ21,"
    Print #1, "vert2 = SRX22, -SRY2-RBt, SRZ22,"
    Print #1, "vert3 = SRX23, -SRY2-RBt, SRZ23,"
    Print #1, "vert4 = SRX24, -SRY2-RBt, SRZ24,"
    Print #1, "vert5 = SRX25, -SRY2-RBt, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX31, SRY3, SRZ31,"
    Print #1, "vert2 = SRX32, SRY3, SRZ32,"
    Print #1, "vert3 = SRX33, SRY3, SRZ33,"
    Print #1, "vert4 = SRX34, SRY3, SRZ34,"
    Print #1, "vert5 = SRX35, SRY3, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX31, -SRY3, SRZ31,"
    Print #1, "vert2 = -SRX32, -SRY3, SRZ32,"
    Print #1, "vert3 = -SRX33, -SRY3, SRZ33,"
    Print #1, "vert4 = -SRX34, -SRY3, SRZ34,"
    Print #1, "vert5 = -SRX35, -SRY3, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBt, SRZ11,"
    Print #1, "vert2 = SRX11, SRY1+RBt, SRZ14,"
    Print #1, "vert3 = SRX11-HBd, SRY1+RBt, SRZ14,"
    Print #1, "vert4 = SRX11-HBd, SRY1+RBt, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1, SRZ11,"
    Print #1, "vert2 = SRX11, -SRY1, SRZ14,"
    Print #1, "vert3 = SRX11-HBd, -SRY1, SRZ14,"
    Print #1, "vert4 = SRX11-HBd, -SRY1, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type03_Y(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------"

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"
    Print #1, "assign DimG = HBw/2, var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal"
    Print #1, "assign WRX11 =  (HBw/2) , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX12 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX13 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX14 =  WRX11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX15 =  (HBw/2), var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRX21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX22 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX23 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX24 =  -WRX11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY2 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRY3 =  RBtH , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal"
    Print #1, "assign SRY11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY12 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY13 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY14 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRY21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY22 =  -SRY11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY23 =  -SRY11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY24 =  -SRY11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX2 =  DimG , var_type=""Float"";"

    Print #1, "assign SRY31 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ31 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY32 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ32 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY33 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ33 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY34 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ34 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY35 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ35 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX3 =  -RBtH, var_type=""Float"";"

    ' !BASE PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = WRX11+RBt, WRY1, WRZ11,"
    Print #1, "vert2 = WRX12, WRY1, WRZ12,"
    Print #1, "vert3 = WRX13, WRY1, WRZ13,"
    Print #1, "vert4 = WRX14+RBt, WRY1, WRZ14,"
    Print #1, "vert5 = WRX15+RBt, WRY1, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21-RBt, WRY2-RBt, WRZ21,"
    Print #1, "vert2 = WRX22, WRY2-RBt, WRZ22,"
    Print #1, "vert3 = WRX23, WRY2-RBt, WRZ23,"
    Print #1, "vert4 = WRX24-RBt, WRY2-RBt, WRZ24,"
    Print #1, "vert5 = WRX25-RBt, WRY2-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21-RBt, -WRY2, WRZ21,"
    Print #1, "vert2 = WRX22, -WRY2, WRZ22,"
    Print #1, "vert3 = WRX23, -WRY2, WRZ23,"
    Print #1, "vert4 = WRX24-RBt, -WRY2, WRZ24,"
    Print #1, "vert5 = WRX25-RBt, -WRY2, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX11+RBt, -WRY1+RBt, WRZ11,"
    Print #1, "vert2 = WRX12, -WRY1+RBt, WRZ12,"
    Print #1, "vert3 = WRX13, -WRY1+RBt, WRZ13,"
    Print #1, "vert4 = WRX14+RBt, -WRY1+RBt, WRZ14,"
    Print #1, "vert5 = WRX15+RBt, -WRY1+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX11+RBt, WRY3, WRZ11,"
    Print #1, "vert2 = WRX12, WRY3, WRZ12,"
    Print #1, "vert3 = WRX13, WRY3, WRZ13,"
    Print #1, "vert4 = WRX14+RBt, WRY3, WRZ14,"
    Print #1, "vert5 = WRX15+RBt, WRY3, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX11-RBt, -WRY3, WRZ11,"
    Print #1, "vert2 = -WRX12, -WRY3, WRZ12,"
    Print #1, "vert3 = -WRX13, -WRY3, WRZ13,"
    Print #1, "vert4 = -WRX14-RBt, -WRY3, WRZ14,"
    Print #1, "vert5 = -WRX15-RBt, -WRY3, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' !Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX1, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1, SRY12, SRZ12,"
    Print #1, "vert3 = SRX1, SRY13, SRZ13,"
    Print #1, "vert4 = SRX1, SRY14, SRZ14,"
    Print #1, "vert5 = SRX1, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX2+RBt, SRY21, SRZ21,"
    Print #1, "vert2 = SRX2+RBt, SRY22, SRZ22,"
    Print #1, "vert3 = SRX2+RBt, SRY23, SRZ23,"
    Print #1, "vert4 = SRX2+RBt, SRY24, SRZ24,"
    Print #1, "vert5 = SRX2+RBt, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBt, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBt, SRY12, SRZ12,"
    Print #1, "vert3 = -SRX1-RBt, SRY13, SRZ13,"
    Print #1, "vert4 = -SRX1-RBt, SRY14, SRZ14,"
    Print #1, "vert5 = -SRX1-RBt, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX2, SRY21, SRZ21,"
    Print #1, "vert2 = -SRX2, SRY22, SRZ22,"
    Print #1, "vert3 = -SRX2, SRY23, SRZ23,"
    Print #1, "vert4 = -SRX2, SRY24, SRZ24,"
    Print #1, "vert5 = -SRX2, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX3, SRY31, SRZ31,"
    Print #1, "vert2 = SRX3, SRY32, SRZ32,"
    Print #1, "vert3 = SRX3, SRY33, SRZ33,"
    Print #1, "vert4 = SRX3, SRY34, SRZ34,"
    Print #1, "vert5 = SRX3, SRY35, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX3, -SRY31, SRZ31,"
    Print #1, "vert2 = -SRX3, -SRY32, SRZ32,"
    Print #1, "vert3 = -SRX3, -SRY33, SRZ33,"
    Print #1, "vert4 = -SRX3, -SRY34, SRZ34,"
    Print #1, "vert5 = -SRX3, -SRY35, SRZ35,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX1, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1, SRY11, SRZ14,"
    Print #1, "vert3 = SRX1, SRY11-HBd, SRZ14,"
    Print #1, "vert4 = SRX1, SRY11-HBd, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBt, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBt, SRY11, SRZ14,"
    Print #1, "vert3 = -SRX1-RBt, SRY11-HBd, SRZ14,"
    Print #1, "vert4 = -SRX1-RBt, SRY11-HBd, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type04_X(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"
    Print #1, "assign DimG = HBw/2, var_type=""Float"";"


    ' !Weak Axis Rib Plate Point Cal
    Print #1, "assign WRY11 =  HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY12 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY13 =  WRY11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY14 =  WRY11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY15 =  HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRY21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY22 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRY23 =  -WRY11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRY24 =  -WRY11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal
    Print #1, "assign SRX11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX12 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX13 =  SRX11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX14 =  SRX11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRX21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX22 =  -SRX11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRX23 =  -SRX11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRX24 =  -SRX11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY2 =  DimG , var_type=""Float"";"

    ' !BASE PLATE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = WRX1-RBt, WRY11+RBt, WRZ11,"
    Print #1, "vert2 = WRX1-RBt, WRY12, WRZ12,"
    Print #1, "vert3 = WRX1-RBt, WRY13, WRZ13,"
    Print #1, "vert4 = WRX1-RBt, WRY14+RBt, WRZ14,"
    Print #1, "vert5 = WRX1-RBt, WRY15+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX2, WRY21-RBt, WRZ21,"
    Print #1, "vert2 = WRX2, WRY22, WRZ22,"
    Print #1, "vert3 = WRX2, WRY23, WRZ23,"
    Print #1, "vert4 = WRX2, WRY24-RBt, WRZ24,"
    Print #1, "vert5 = WRX2, WRY25-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX2+RBt, WRY21-RBt, WRZ21,"
    Print #1, "vert2 = -WRX2+RBt, WRY22, WRZ22,"
    Print #1, "vert3 = -WRX2+RBt, WRY23, WRZ23,"
    Print #1, "vert4 = -WRX2+RBt, WRY24-RBt, WRZ24,"
    Print #1, "vert5 = -WRX2+RBt, WRY25-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -WRX1, WRY11+RBt, WRZ11,"
    Print #1, "vert2 = -WRX1, WRY12, WRZ12,"
    Print #1, "vert3 = -WRX1, WRY13, WRZ13,"
    Print #1, "vert4 = -WRX1, WRY14+RBt, WRZ14,"
    Print #1, "vert5 = -WRX1, WRY15+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' !Strong Axis RIB PLAPTE MODELING
    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBt, SRZ11,"
    Print #1, "vert2 = SRX12, SRY1+RBt, SRZ12,"
    Print #1, "vert3 = SRX13, SRY1+RBt, SRZ13,"
    Print #1, "vert4 = SRX14, SRY1+RBt, SRZ14,"
    Print #1, "vert5 = SRX15, SRY1+RBt, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, SRY2, SRZ21,"
    Print #1, "vert2 = SRX22, SRY2, SRZ22,"
    Print #1, "vert3 = SRX23, SRY2, SRZ23,"
    Print #1, "vert4 = SRX24, SRY2, SRZ24,"
    Print #1, "vert5 = SRX25, SRY2, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1, SRZ11,"
    Print #1, "vert2 = SRX12, -SRY1, SRZ12,"
    Print #1, "vert3 = SRX13, -SRY1, SRZ13,"
    Print #1, "vert4 = SRX14, -SRY1, SRZ14,"
    Print #1, "vert5 = SRX15, -SRY1, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX21, -SRY2-RBt, SRZ21,"
    Print #1, "vert2 = SRX22, -SRY2-RBt, SRZ22,"
    Print #1, "vert3 = SRX23, -SRY2-RBt, SRZ23,"
    Print #1, "vert4 = SRX24, -SRY2-RBt, SRZ24,"
    Print #1, "vert5 = SRX25, -SRY2-RBt, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, SRY1+RBt, SRZ11,"
    Print #1, "vert2 = SRX11, SRY1+RBt, SRZ14,"
    Print #1, "vert3 = SRX11-HBd, SRY1+RBt, SRZ14,"
    Print #1, "vert4 = SRX11-HBd, SRY1+RBt, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX11, -SRY1, SRZ11,"
    Print #1, "vert2 = SRX11, -SRY1, SRZ14,"
    Print #1, "vert3 = SRX11-HBd, -SRY1, SRZ14,"
    Print #1, "vert4 = SRX11-HBd, -SRY1, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type04_Y(subPath As String, _
    subCdepth As Single, subCwidth As Single, subCwt As Single, subCft As Single, _
    subLx As Single, subLy As Single, subLz As Single, _
    subRBt As Single, subRBr As Single, subRBe As Single, _
    subRBsw As Single, subRBww As Single, subRBh As Single)
    
Open subPath For Append As #1

    ' !------------------ Data Input Start ------------------------

    Print #1, "assign HBd = " & CStr(subCdepth) & ", var_type=""Float"";"
    Print #1, "assign HBw = " & CStr(subCwidth) & ", var_type=""Float"";"
    Print #1, "assign HBwt = " & CStr(subCwt) & ", var_type=""Float"";"
    Print #1, "assign HBft = " & CStr(subCft) & ", var_type=""Float"";"

    Print #1, "assign Lx = " & CStr(subLx) & ", var_type=""Float"";"
    Print #1, "assign Ly = " & CStr(subLy) & ", var_type=""Float"";"
    Print #1, "assign Lz = " & CStr(subLz) & ", var_type=""Float"";"

    Print #1, "assign RBt = " & CStr(subRBt) & ", var_type=""Float"";"
    Print #1, "assign RBr = " & CStr(subRBr) & ", var_type=""Float"";"
    Print #1, "assign RBe = " & CStr(subRBe) & ", var_type=""Float"";"

    ' !------------------ Data Input End ------------------------"

    Print #1, "assign RBsw = " & CStr(subRBsw) & ", var_type=""Float"";"
    Print #1, "assign RBww = " & CStr(subRBww) & ", var_type=""Float"";"
    Print #1, "assign RBh = " & CStr(subRBh) & ", var_type=""Float"";"

    Print #1, "assign Lhx = " & CStr(subLx / 2) & ", var_type=""Float"";"
    Print #1, "assign Lhy = " & CStr(subLy / 2) & ", var_type=""Float"";"

    Print #1, "assign RBtH = " & CStr(subRBt / 2) & ", var_type=""Float"";"
    Print #1, "assign DimG = HBw/2, var_type=""Float"";"

    ' !Weak Axis Rib Plate Point Cal"
    Print #1, "assign WRX11 =  (HBw/2) , var_type=""Float"";"
    Print #1, "assign WRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX12 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX13 =  WRX11+RBww-RBe , var_type=""Float"";"
    Print #1, "assign WRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX14 =  WRX11+RBr , var_type=""Float"";"
    Print #1, "assign WRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX15 =  (HBw/2), var_type=""Float"";"
    Print #1, "assign WRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY1 =  HBd/2 , var_type=""Float"";"

    Print #1, "assign WRX21 =  -HBw/2 , var_type=""Float"";"
    Print #1, "assign WRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX22 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign WRX23 =  -WRX11-RBww+RBe , var_type=""Float"";"
    Print #1, "assign WRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign WRX24 =  -WRX11-RBr , var_type=""Float"";"
    Print #1, "assign WRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRX25 =  -HBw/2, var_type=""Float"";"
    Print #1, "assign WRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign WRY2 =  HBd/2 , var_type=""Float"";"

    ' !Strong Axis Rib Plate Point Cal"
    Print #1, "assign SRY11 =  HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ11 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY12 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ12 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY13 =  SRY11+RBsw-RBe , var_type=""Float"";"
    Print #1, "assign SRZ13 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY14 =  SRY11+RBr , var_type=""Float"";"
    Print #1, "assign SRZ14 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY15 =  HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ15 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX1 =  DimG , var_type=""Float"";"

    Print #1, "assign SRY21 =  -HBd/2 , var_type=""Float"";"
    Print #1, "assign SRZ21 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY22 =  -SRY11-RBSw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ22 =  Lz, var_type=""Float"";"
    Print #1, "assign SRY23 =  -SRY11-RBsw+RBe , var_type=""Float"";"
    Print #1, "assign SRZ23 =  Lz+RBr , var_type=""Float"";"
    Print #1, "assign SRY24 =  -SRY11-RBr , var_type=""Float"";"
    Print #1, "assign SRZ24 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRY25 =  -HBd/2, var_type=""Float"";"
    Print #1, "assign SRZ25 =  Lz+RBh , var_type=""Float"";"
    Print #1, "assign SRX2 =  DimG , var_type=""Float"";"

    ' !BASE PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Lhx, Lhy, 0,"
    Print #1, "vert2 = -Lhx, Lhy, 0,"
    Print #1, "vert3 = -Lhx, -Lhy, 0,"
    Print #1, "vert4 = Lhx,-Lhy, 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(subLz) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = Lz;"

    ' !Weak Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = WRX11+RBt, WRY1, WRZ11,"
    Print #1, "vert2 = WRX12, WRY1, WRZ12,"
    Print #1, "vert3 = WRX13, WRY1, WRZ13,"
    Print #1, "vert4 = WRX14+RBt, WRY1, WRZ14,"
    Print #1, "vert5 = WRX15+RBt, WRY1, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21-RBt, WRY2-RBt, WRZ21,"
    Print #1, "vert2 = WRX22, WRY2-RBt, WRZ22,"
    Print #1, "vert3 = WRX23, WRY2-RBt, WRZ23,"
    Print #1, "vert4 = WRX24-RBt, WRY2-RBt, WRZ24,"
    Print #1, "vert5 = WRX25-RBt, WRY2-RBt, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX21-RBt, -WRY2, WRZ21,"
    Print #1, "vert2 = WRX22, -WRY2, WRZ22,"
    Print #1, "vert3 = WRX23, -WRY2, WRZ23,"
    Print #1, "vert4 = WRX24-RBt, -WRY2, WRZ24,"
    Print #1, "vert5 = WRX25-RBt, -WRY2, WRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = WRX11+RBt, -WRY1+RBt, WRZ11,"
    Print #1, "vert2 = WRX12, -WRY1+RBt, WRZ12,"
    Print #1, "vert3 = WRX13, -WRY1+RBt, WRZ13,"
    Print #1, "vert4 = WRX14+RBt, -WRY1+RBt, WRZ14,"
    Print #1, "vert5 = WRX15+RBt, -WRY1+RBt, WRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    ' !Strong Axis RIB PLAPTE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = SRX1, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1, SRY12, SRZ12,"
    Print #1, "vert3 = SRX1, SRY13, SRZ13,"
    Print #1, "vert4 = SRX1, SRY14, SRZ14,"
    Print #1, "vert5 = SRX1, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX2+RBt, SRY21, SRZ21,"
    Print #1, "vert2 = SRX2+RBt, SRY22, SRZ22,"
    Print #1, "vert3 = SRX2+RBt, SRY23, SRZ23,"
    Print #1, "vert4 = SRX2+RBt, SRY24, SRZ24,"
    Print #1, "vert5 = SRX2+RBt, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBt, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBt, SRY12, SRZ12,"
    Print #1, "vert3 = -SRX1-RBt, SRY13, SRZ13,"
    Print #1, "vert4 = -SRX1-RBt, SRY14, SRZ14,"
    Print #1, "vert5 = -SRX1-RBt, SRY15, SRZ15,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX2, SRY21, SRZ21,"
    Print #1, "vert2 = -SRX2, SRY22, SRZ22,"
    Print #1, "vert3 = -SRX2, SRY23, SRZ23,"
    Print #1, "vert4 = -SRX2, SRY24, SRZ24,"
    Print #1, "vert5 = -SRX2, SRY25, SRZ25,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = SRX1, SRY11, SRZ11,"
    Print #1, "vert2 = SRX1, SRY11, SRZ14,"
    Print #1, "vert3 = SRX1, SRY11-HBd, SRZ14,"
    Print #1, "vert4 = SRX1, SRY11-HBd, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -SRX1-RBt, SRY11, SRZ11,"
    Print #1, "vert2 = -SRX1-RBt, SRY11, SRZ14,"
    Print #1, "vert3 = -SRX1-RBt, SRY11-HBd, SRZ14,"
    Print #1, "vert4 = -SRX1-RBt, SRY11-HBd, SRZ11,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""RP_" & CStr(subRBt) & """" & ", " & _
              "ov_parall = 1" & ", "
    Print #1, "thickness = RBt;"
Close #1
End Sub

Public Sub BP_Fixed_PML_Type06(subPath As String, _
    subLx As Single, subLy As Single, subLz As Single)
    
Dim TempX1 As Single, TempX2 As Single, TempY1 As Single, TempY2 As Single, TempThk As Single

TempX1 = subLx / 2
TempX2 = -(subLy / 2)
TempY1 = subLx / 2
TempY2 = -(subLy / 2)
TempThk = subLz
Call Print_RecBox(subPath, TempX1, TempX2, TempY1, TempY2, TempThk, "BP_" & CStr(TempThk))

End Sub

Public Sub BP_Fixed_PML_Nut(valPathName As String, valPt As Single, _
    valXbtob As Single, valYbtob As Single, _
    valBoltDia As Single, valNutDia As Single, valNutHei As Single, valBoltName As String, _
    valType As String, valDir As String)

Dim Seta(1 To 12) As Single
Dim Nx() As Single, Ny() As Single
Dim BX(1 To 12) As Single, By(1 To 12) As Single
Dim valBoltEA As Integer, i As Integer, j As Integer

For i = 1 To 6
    If i = 1 Then
        Seta(1) = 60
    Else
        Seta(i) = 60 + Seta(i - 1)
    End If
Next i
    
If valDir = "VectorY" Then
    Select Case valType
        Case "Type01", "Type05"
            valBoltEA = 6
            
            BX(1) = -valXbtob
            BX(2) = 0
            BX(3) = valXbtob
            By(1) = valYbtob
            By(2) = valYbtob
            By(3) = valYbtob
            
            BX(4) = -valXbtob
            BX(5) = 0
            BX(6) = valXbtob
            By(4) = -valYbtob
            By(5) = -valYbtob
            By(6) = -valYbtob
        Case "Type02"
            valBoltEA = 8
            
            BX(1) = -valXbtob - (valXbtob / 2)
            BX(2) = -(valXbtob / 2)
            BX(3) = valXbtob / 2
            BX(4) = valXbtob + (valXbtob / 2)
            By(1) = valYbtob
            By(2) = valYbtob
            By(3) = valYbtob
            By(4) = valYbtob
            
            BX(5) = -valXbtob - (valXbtob / 2)
            BX(6) = -(valXbtob / 2)
            BX(7) = valXbtob / 2
            BX(8) = valXbtob + (valXbtob / 2)
            By(5) = -valYbtob
            By(6) = -valYbtob
            By(7) = -valYbtob
            By(8) = -valYbtob
        Case "Type03"
            valBoltEA = 12
            
            BX(1) = -valXbtob - (valXbtob / 2)
            BX(2) = -(valXbtob / 2)
            BX(3) = valXbtob / 2
            BX(4) = valXbtob + (valXbtob / 2)
            By(1) = valYbtob + valYbtob / 2
            By(2) = valYbtob + valYbtob / 2
            By(3) = valYbtob + valYbtob / 2
            By(4) = valYbtob + valYbtob / 2
            
            BX(5) = -valXbtob - (valXbtob / 2)
            BX(6) = -(valXbtob / 2)
            BX(7) = valXbtob / 2
            BX(8) = valXbtob + (valXbtob / 2)
            By(5) = -valYbtob - valYbtob / 2
            By(6) = -valYbtob - valYbtob / 2
            By(7) = -valYbtob - valYbtob / 2
            By(8) = -valYbtob - valYbtob / 2
            
            BX(9) = -valXbtob - (valXbtob / 2)
            BX(10) = -valXbtob - (valXbtob / 2)
            BX(11) = valXbtob + (valXbtob / 2)
            BX(12) = valXbtob + (valXbtob / 2)
            By(9) = -valYbtob / 2
            By(10) = valYbtob / 2
            By(11) = -valYbtob / 2
            By(12) = valYbtob / 2
        Case "Type04"
            valBoltEA = 8
            
            BX(1) = -valXbtob
            BX(2) = 0
            BX(3) = valXbtob
            By(1) = valYbtob
            By(2) = valYbtob
            By(3) = valYbtob
            
            BX(4) = -valXbtob
            BX(5) = 0
            BX(6) = valXbtob
            By(4) = -valYbtob
            By(5) = -valYbtob
            By(6) = -valYbtob
            
            BX(7) = -valXbtob
            BX(8) = valXbtob
            By(7) = 0
            By(8) = 0
        Case "Type06"
            valBoltEA = 6
            
            BX(1) = -valXbtob
            BX(2) = -valXbtob
            BX(3) = -valXbtob
            By(1) = -valYbtob
            By(2) = 0
            By(3) = valYbtob
            
            BX(4) = valXbtob
            BX(5) = valXbtob
            BX(6) = valXbtob
            By(4) = -valYbtob
            By(5) = 0
            By(6) = valYbtob
        Case "Type07"
            valBoltEA = 4
            
            BX(1) = -valXbtob
            BX(2) = valXbtob
            By(1) = valYbtob
            By(2) = valYbtob
            
            BX(3) = -valXbtob
            BX(4) = valXbtob
            By(3) = -valYbtob
            By(4) = -valYbtob
    End Select
Else
    Select Case valType
        Case "Type01", "Type05"
            valBoltEA = 6
            
            BX(1) = valXbtob
            BX(2) = valXbtob
            BX(3) = valXbtob
            By(1) = -valYbtob
            By(2) = 0
            By(3) = valYbtob
            
            BX(4) = -valXbtob
            BX(5) = -valXbtob
            BX(6) = -valXbtob
            By(4) = -valYbtob
            By(5) = 0
            By(6) = valYbtob
        Case "Type02"
            valBoltEA = 8
            
            BX(1) = valXbtob
            BX(2) = valXbtob
            BX(3) = valXbtob
            BX(4) = valXbtob
            By(1) = -valYbtob - (valYbtob / 2)
            By(2) = -(valYbtob / 2)
            By(3) = valYbtob / 2
            By(4) = valYbtob + (valYbtob / 2)
            
            BX(5) = -valXbtob
            BX(6) = -valXbtob
            BX(7) = -valXbtob
            BX(8) = -valXbtob
            By(5) = -valYbtob - (valYbtob / 2)
            By(6) = -(valYbtob / 2)
            By(7) = valYbtob / 2
            By(8) = valYbtob + (valYbtob / 2)
        Case "Type03"
            valBoltEA = 12
            
            BX(1) = valXbtob + valXbtob / 2
            BX(2) = valXbtob + valXbtob / 2
            BX(3) = valXbtob + valXbtob / 2
            BX(4) = valXbtob + valXbtob / 2
            By(1) = -valYbtob - (valYbtob / 2)
            By(2) = -(valYbtob / 2)
            By(3) = valYbtob / 2
            By(4) = valYbtob + (valYbtob / 2)
            
            BX(5) = -valXbtob - valXbtob / 2
            BX(6) = -valXbtob - valXbtob / 2
            BX(7) = -valXbtob - valXbtob / 2
            BX(8) = -valXbtob - valXbtob / 2
            By(5) = -valYbtob - (valYbtob / 2)
            By(6) = -(valYbtob / 2)
            By(7) = valYbtob / 2
            By(8) = valYbtob + (valYbtob / 2)
            
            BX(9) = -valXbtob / 2
            BX(10) = valXbtob / 2
            BX(11) = -valXbtob / 2
            BX(12) = valXbtob / 2
            By(9) = -valYbtob - (valYbtob / 2)
            By(10) = -valYbtob - (valYbtob / 2)
            By(11) = valYbtob + (valYbtob / 2)
            By(12) = valYbtob + (valYbtob / 2)
        Case "Type04"
            valBoltEA = 8
            
            BX(1) = valXbtob
            BX(2) = valXbtob
            BX(3) = valXbtob
            By(1) = -valYbtob
            By(2) = 0
            By(3) = valYbtob
            
            BX(4) = -valXbtob
            BX(5) = -valXbtob
            BX(6) = -valXbtob
            By(4) = -valYbtob
            By(5) = 0
            By(6) = valYbtob
            
            BX(7) = 0
            BX(8) = 0
            By(7) = -valYbtob
            By(8) = valYbtob
        Case "Type06"
            valBoltEA = 6
            
            BX(1) = -valXbtob
            BX(2) = 0
            BX(3) = valXbtob
            By(1) = -valYbtob
            By(2) = -valYbtob
            By(3) = -valYbtob
            
            BX(4) = -valXbtob
            BX(5) = 0
            BX(6) = valXbtob
            By(4) = valYbtob
            By(5) = valYbtob
            By(6) = valYbtob
        Case "Type07"
            valBoltEA = 4
            
            BX(1) = valXbtob
            BX(2) = valXbtob
            By(1) = -valYbtob
            By(2) = valYbtob
            
            BX(3) = -valXbtob
            BX(4) = -valXbtob
            By(3) = -valYbtob
            By(4) = valYbtob
    End Select
End If

ReDim Nx(1 To valBoltEA, 1 To 6) As Single
ReDim Ny(1 To valBoltEA, 1 To 6) As Single

For i = 1 To valBoltEA
    For j = 1 To 6
        Nx(i, j) = (BX(i) + (valNutDia / 2) * Sin((Seta(j) / 180) * 3.14))
        Ny(i, j) = (By(i) + (valNutDia / 2) * Cos((Seta(j) / 180) * 3.14))
    Next j
Next i


Open valPathName For Append As #1
    For i = 1 To valBoltEA

        Print #1, "plc_area"
        For j = 1 To 6
            Print #1, "vert" & j & " = " & Format(Nx(i, j), "0.000") & ", " & _
                                           Format(Ny(i, j), "0.000") & "," & _
                                           Format(valPt + valNutHei, "0.000") & ","
        Next j
        
        Print #1, "class = " & gstr_BPClass & ", " & _
                  "grade = """ & gstr_Grade & """, " & _
                  "material = """ & gstr_Material & """, " & _
                  "name = ""AB_" & CStr(valBoltName) & """" & ", " & _
                  "ov_parall = 1" & ", "
        Print #1, "thickness = " & valNutHei & ";"
    Next i
    
Close #1

End Sub

Public Function gf_StringtoSingle(as_Value As String) As Single
    If Not IsNull(as_Value) Then
        If Trim(as_Value) = "" Then
            gf_StringtoSingle = 0
        Else
            If IsNumeric(Trim(as_Value)) Then
                gf_StringtoSingle = CSng(Trim(as_Value))
            Else
                gf_StringtoSingle = 0
            End If
        End If
    Else
        gf_StringtoSingle = 0
    End If
End Function
Public Sub gs_CobAddItem(valCobList As ComboBox)

Dim sql As String
valCobList.Clear

sql = "Select * from code"

Set adoRecordset = adoConnection.Execute(sql)

Do Until adoRecordset.EOF
    valCobList.AddItem adoRecordset!CodeName
    adoRecordset.MoveNext
Loop

adoRecordset.Close

Set adoRecordset = Nothing


'valCobList.AddItem "JIS"
'valCobList.AddItem "AISC"

End Sub
Public Sub gs_project_Output(ByVal valProjectName As String)
Open App.Path & "\Project.ini" For Output As #1
            Print #1, valProjectName
Close #1

End Sub
Public Sub gs_project_Input(ByRef valProjectName As String)
Open App.Path & "\Project.ini" For Input As #1
            Input #1, valProjectName
Close #1

End Sub
Public Sub gs_Call_Project(ByRef refProjectName As String, ByRef refError As Integer)

On Error GoTo ErrorLabel
refError = 0
Open App.Path & "\Project.ini" For Input As #1
            Input #1, refProjectName
Close #1
Exit Sub
ErrorLabel:
refError = 1
End Sub

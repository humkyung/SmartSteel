Attribute VB_Name = "mdlHorBracing"

Public Sub HB_PML(valPathName As String, ByVal valJobName As String, ByVal valCode As String, _
    ByVal valFormCode As String, valHBGP_Flag As String, valType As String, valUnit As String, _
    valBeam As String, valSubBeam As String, valBracing As String, _
    valSpace3 As Single, valNutFlag As Integer)

Dim TempSubBedepth As Single, TempSubBewidth As Single, TempSubBewt As Single, TempSubBeft As Single
Dim TempBedepth As Single, TempBewidth As Single, TempBewt As Single
Dim TempBdepth As Single, TempBwidth As Single, TempBwt As Single
Dim TempSP1 As Single, TempSP2 As Single, TempSP3 As Single
Dim TempHTB_Space As Single, TempGPThk As Single, TempOF As Single
Dim TempBoltName As String, TempNutDia As Single, TempNutHei As Single, TempGage As Single

    Call HBData_Call(valJobName, valCode, valFormCode, valBeam, valSubBeam, valBracing)

    TempSubBedepth = convert(gsin_SubBedepth, "mm", valUnit)
    TempSubBewidth = convert(gsin_SubBewidth, "mm", valUnit)
    TempSubBeft = convert(gsin_SubBeft, "mm", valUnit)
    TempBedepth = convert(gsin_Bedepth, "mm", valUnit)
    TempBewidth = convert(gsin_Bewidth, "mm", valUnit)
    TempBewt = convert(gsin_Bewt, "mm", valUnit)
    TempBdepth = convert(gsin_Bdepth, "mm", valUnit)
    TempBwidth = convert(gsin_Bwidth, "mm", valUnit)
    TempBwt = convert(gsin_Bwt, "mm", valUnit)
    
    TempSP1 = convert(gsin_SP1, gstr_Unit, valUnit)
    TempSP2 = convert(gsin_SP2, gstr_Unit, valUnit)
    TempSP3 = valSpace3
    TempHTB_Space = convert(gsin_HTB_Space, gstr_Unit, valUnit)
    TempGPThk = convert(gsin_GPThk, gstr_Unit, valUnit)
    TempGage = convert(gsin_Gage, gstr_Unit, valUnit)
    
    TempBoltName = gstr_BoltName
    TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
    TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)
    
    Select Case valHBGP_Flag
        Case "Module01"
            Select Case valType
                Case "Type-01"
                    Call HB_WEAK_LeftBott_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-03"
                    Call HB_WEAK_LeftBott_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-02"
                    Call HB_WEAK_LeftBott_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-04"
                    Call HB_STRONG_LeftBott_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-06"
                    Call HB_STRONG_LeftBott_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-05"
                    Call HB_STRONG_LeftBott_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
            End Select
               If valNutFlag = 1 Then
                              Call Hor_HTB_Nut_M01(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
               End If
        Case "Module02"
            Select Case valType
                Case "Type-01"
                    Call HB_WEAK_RightBott_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-03"
                    Call HB_WEAK_RightBott_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-02"
                    Call HB_WEAK_RightBott_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-04"
                    Call HB_STRONG_RightBott_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-06"
                    Call HB_STRONG_RightBott_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-05"
                    Call HB_STRONG_RightBott_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
            End Select
               If valNutFlag = 1 Then
                              Call Hor_HTB_Nut_M02(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                               TempHTB_Space, TempGPThk, TempGage, TempBwt)
               End If
        Case "Module03"
            Select Case valType
                Case "Type-01"
                    Call HB_WEAK_RightTop_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-03"
                    Call HB_WEAK_RightTop_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-02"
                    Call HB_WEAK_RightTop_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-04"
                    Call HB_STRONG_RightTop_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-06"
                    Call HB_STRONG_RightTop_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-05"
                    Call HB_STRONG_RightTop_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
            End Select
            If valNutFlag = 1 Then Call Hor_HTB_Nut_M03(valPathName, gstr_BoltType, _
                                                                                                       TempBoltName, TempNutDia, _
                                                                                                       TempNutHei, gsin_HTB_Num, _
                                                                                                       TempHTB_Space, TempGPThk, _
                                                                                                       TempGage, TempBwt)
        Case "Module04"
            Select Case valType
                Case "Type-01"
                    Call HB_WEAK_LeftTop_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-03"
                    Call HB_WEAK_LeftTop_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-02"
                    Call HB_WEAK_LeftTop_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-04"
                    Call HB_STRONG_LeftTop_Center(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-06"
                    Call HB_STRONG_LeftTop_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                Case "Type-05"
                    Call HB_STRONG_LeftTop_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                        TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
            End Select
            If valNutFlag = 1 Then Call Hor_HTB_Nut_M04(valPathName, gstr_BoltType, TempBoltName, _
                                                                                                         TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                                                                         TempHTB_Space, TempGPThk, TempGage, TempBwt)
        Case "Module05"
            Select Case valType
                Case "Type-01"
                              Call HB_XBracing_Type01(valPathName, TempBwidth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then Call Hor_HTB_Nut_M05_1(valPathName, gstr_BoltType, TempBoltName, _
                                                                                                                                       TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                                                                                                       TempHTB_Space, TempGPThk, TempGage, TempBwt)
                Case "Type-02"
                              Call HB_XBracing_Type02(valPathName, TempBwidth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then Call Hor_HTB_Nut_M05_2(valPathName, gstr_BoltType, TempBoltName, _
                                                                                                                                       TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                                                                                                       TempHTB_Space, TempGPThk, TempGage, TempBwt)
                Case "Type-03"
                              Call HB_XBracing_Type03(valPathName, TempBwidth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then Call Hor_HTB_Nut_M05_3(valPathName, gstr_BoltType, TempBoltName, _
                                                                                                                                       TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                                                                                                       TempHTB_Space, TempGPThk, TempGage, TempBwt)
            End Select
        Case "Module06"
            Select Case valType
                Case "Type-01"
                              Call HB_KBracing_Left(valPathName, TempBewidth, TempBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then Call Hor_HTB_Nut_M06_a(valPathName, gstr_BoltType, TempBoltName, _
                                                                                                                                       TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                                                                                                       TempHTB_Space, TempGPThk, TempGage, TempBwt)
                Case "Type-02"
                              Call HB_KBracing_Bottom(valPathName, TempBewidth, TempBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M06_b1(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
                                             Call Hor_HTB_Nut_M06_b2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
                Case "Type-03"
                              Call HB_KBracing_Right(valPathName, TempBewidth, TempBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M06_c(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
                Case "Type-04"
                              Call HB_KBracing_Top(valPathName, TempBewidth, TempBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M06_d1(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
                                             Call Hor_HTB_Nut_M06_d2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, gsin_HTB_Num, _
                                                            TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
            End Select
        Case "Module07"
            Select Case valType
                Case "Type-01"
                              Call HB_BtoB_LeftBottom_Center(valPathName, TempBewidth, TempBewt, TempSubBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M01(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                                                          TempNutHei, gsin_HTB_Num, _
                                                                                          TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
                Case "Type-02"
                              Call HB_BtoB_RightBottom_Center(valPathName, TempBewidth, TempBewt, TempSubBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M02(valPathName, gstr_BoltType, TempBoltName, TempNutDia, TempNutHei, _
                                                                                                         gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
                Case "Type-03"
                              Call HB_BtoB_RightTop_Center(valPathName, TempBewidth, TempBewt, TempSubBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M03(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                                                                         TempNutHei, gsin_HTB_Num, _
                                                                                                         TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
                Case "Type-04"
                              Call HB_BtoB_LeftTop_Center(valPathName, TempBewidth, TempBewt, TempSubBewt, TempBwidth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk)
                              If valNutFlag = 1 Then
                                             Call Hor_HTB_Nut_M04(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                                                          TempNutHei, gsin_HTB_Num, _
                                                                                          TempHTB_Space, TempGPThk, TempGage, TempBwt)
                              End If
            End Select
    End Select
End Sub

Public Sub HB_STRONG_LeftBott_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, GPT,"
    Print #1, "vert2 = p21, p22, GPT,"
    Print #1, "vert3 = p31, p32, GPT,"
    Print #1, "vert4 = p41, P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_STRONG_LeftBott_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, " evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W/cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al+move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al+move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al+move, var_type = ""float"";"

    Print #1, "assign TempX = (p31-BewtH)*tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32-TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, GPT,"
    Print #1, "vert2 = p21, p22, GPT,"
    Print #1, "vert3 = p31, p32, GPT,"
    Print #1, "vert4 = p41, P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_STRONG_LeftBott_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W / Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, GPT,"
    Print #1, "vert2 = p21, p22, GPT,"
    Print #1, "vert3 = p31, p32, GPT,"
    Print #1, "vert4 = p41, P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_LeftBott_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)

Open valPathName For Output As #1
    ' evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/Sin(alpha)) + (0.015/Sin(alpha)) + (W/Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH)/Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, 0,"
    Print #1, "vert2 = p21, p22, 0,"
    Print #1, "vert3 = p31, p32, 0,"
    Print #1, "vert4 = p41, p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_WEAK_LeftBott_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign MoveL = W / Sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, 0,"
    Print #1, "vert2 = p21, p22, 0,"
    Print #1, "vert3 = p31, p32, 0,"
    Print #1, "vert4 = p41, p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_LeftBott_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Move = W / Sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, 0,"
    Print #1, "vert2 = p21, p22, 0,"
    Print #1, "vert3 = p31, p32, 0,"
    Print #1, "vert4 = p41, p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_LeftTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, 0,"
    Print #1, "vert2 = p21, -p22, 0,"
    Print #1, "vert3 = p31, -p32, 0,"
    Print #1, "vert4 = p41, -P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_LeftTop_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W / Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, 0,"
    Print #1, "vert2 = p21, -p22, 0,"
    Print #1, "vert3 = p31, -p32, 0,"
    Print #1, "vert4 = p41, -P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_LeftTop_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W / Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al - Move, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, 0,"
    Print #1, "vert2 = p21, -p22, 0,"
    Print #1, "vert3 = p31, -p32, 0,"
    Print #1, "vert4 = p41, -P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_LeftTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = p31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, GPT,"
    Print #1, "vert2 = p21, -p22, GPT,"
    Print #1, "vert3 = p31, -p32, GPT,"
    Print #1, "vert4 = p41, -p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
    
End Sub

Public Sub HB_WEAK_LeftTop_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign MoveL = W / Sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al - MoveL, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, GPT,"
    Print #1, "vert2 = p21, -p22, GPT,"
    Print #1, "vert3 = p31, -p32, GPT,"
    Print #1, "vert4 = p41, -p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_LeftTop_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign Move = W/sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al+Move, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al+Move, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al+Move, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32-BewtH)/tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, GPT,"
    Print #1, "vert2 = p21, -p22, GPT,"
    Print #1, "vert3 = p31, -p32, GPT,"
    Print #1, "vert4 = p41, -p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightBott_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, 0,"
    Print #1, "vert2 = -p21, p22, 0,"
    Print #1, "vert3 = -p31, p32, 0,"
    Print #1, "vert4 = -p41, P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightBott_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W/cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al-move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al-move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al-move, var_type = ""float"";"

    Print #1, "assign TempX = (p31-BewtH)*tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32-TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, 0,"
    Print #1, "vert2 = -p21, p22, 0,"
    Print #1, "vert3 = -p31, p32, 0,"
    Print #1, "vert4 = -p41, P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightBott_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W / Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, 0,"
    Print #1, "vert2 = -p21, p22, 0,"
    Print #1, "vert3 = -p31, p32, 0,"
    Print #1, "vert4 = -p41, P42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightBott_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32-BewtH)/tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, GPT,"
    Print #1, "vert2 = -p21, p22, GPT,"
    Print #1, "vert3 = -p31, p32, GPT,"
    Print #1, "vert4 = -p41, p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightBott_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Move = W / Sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, GPT,"
    Print #1, "vert2 = -p21, p22, GPT,"
    Print #1, "vert3 = -p31, p32, GPT,"
    Print #1, "vert4 = -p41, p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightBott_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign MoveL = W/sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32-BewtH)/tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, GPT,"
    Print #1, "vert2 = -p21, p22, GPT,"
    Print #1, "vert3 = -p31, p32, GPT,"
    Print #1, "vert4 = -p41, p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, GPT,"
    Print #1, "vert2 = -p21, -p22, GPT,"
    Print #1, "vert3 = -p31, -p32, GPT,"
    Print #1, "vert4 = -p41, -P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightTop_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W/cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al-move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al-move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al-move, var_type = ""float"";"

    Print #1, "assign TempX = (p31-BewtH)*tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32-TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, GPT,"
    Print #1, "vert2 = -p21, -p22, GPT,"
    Print #1, "vert3 = -p31, -p32, GPT,"
    Print #1, "vert4 = -p41, -P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_STRONG_RightTop_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW =" & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"
    Print #1, "assign Move = W / Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al + Move, var_type = ""float"";"

    Print #1, "assign TempX = (p31 - BewtH) * Tan(alpha), var_type = ""float"";"

    Print #1, "assign p41 = BewtH, var_type = ""float"";"
    Print #1, "assign p42 = p32 - TempX, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, GPT,"
    Print #1, "vert2 = -p21, -p22, GPT,"
    Print #1, "vert3 = -p31, -p32, GPT,"
    Print #1, "vert4 = -p41, -P42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32-BewtH)/tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, 0,"
    Print #1, "vert2 = -p21, -p22, 0,"
    Print #1, "vert3 = -p31, -p32, 0,"
    Print #1, "vert4 = -p41, -p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightTop_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD =" & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Move = W / Sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al + Move, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32 - BewtH) / Tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, 0,"
    Print #1, "vert2 = -p21, -p22, 0,"
    Print #1, "vert3 = -p31, -p32, 0,"
    Print #1, "vert4 = -p41, -p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_WEAK_RightTop_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    ' evaluate verify = ""yes"";"

Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    
    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign MoveL = W/sin(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al-MoveL, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign TempX = (p32-BewtH)/tan(alpha), var_type = ""float"";"
    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, 0,"
    Print #1, "vert2 = -p21, -p22, 0,"
    Print #1, "vert3 = -p31, -p32, 0,"
    Print #1, "vert4 = -p41, -p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1

End Sub

Public Sub HB_XBracing_Type01(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Point at left of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign length_x1 = (end2_x - end3_x), var_type = ""float"";"
    Print #1, "assign length_y1 = (end3_y - end2_y), var_type = ""float"";"

    Print #1, "assign alpha = -1*atand(length_y/length_x), var_type = ""float"";"
    'assign alpha1 = atand(length_y1/length_x1), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    'Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    'Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    'Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    'Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type = ""float"";"
    'BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    'BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    'Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type = ""float"";"
    Print #1, "assign alpha1 = atand(W/H), var_type = ""float"";"
    Print #1, "assign Len = sqrt(H**2+W**2), var_type = ""float"";"

    Print #1, "assign Len1 = sqrt(0.01**2+W**2), var_type = ""float"";"
    Print #1, "assign alpha2 = atand(W/0.1), var_type = ""float"";"
    Print #1, "assign alpha3 = 90-alpha2+alpha, var_type = ""float"";"

    Print #1, "assign p11 = len*cos(alpha+alpha1), var_type = ""float"";"
    Print #1, "assign p12 = len*sin(alpha+alpha1), var_type = ""float"";"

    Print #1, "assign p21 = len*cos(alpha-alpha1), var_type = ""float"";"
    Print #1, "assign p22 = len*sin(alpha-alpha1), var_type = ""float"";"

    Print #1, "assign p31 = p21-len1*sin(alpha3), var_type = ""float"";"
    Print #1, "assign p32 = p22-len1*cos(alpha3), var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, GPT,"
    Print #1, "vert2 = p11, p12, GPT,"
    Print #1, "vert3 = p21, p22, GPT,"
    Print #1, "vert4 = -p21, p22, GPT,"
    Print #1, "vert5 = -p11, p12, GPT,"
    Print #1, "vert6 = -p11, -p12, GPT,"
    Print #1, "vert7 = -p21, -p22, GPT,"
    Print #1, "vert8 = p21, -p22, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_XBracing_Type02(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Point at left of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_z = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign length_x1 = (end2_x - end3_x), var_type = ""float"";"
    Print #1, "assign length_z1 = (end3_y - end2_y), var_type = ""float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type = ""float"";"
    'assign alpha1 = atand(length_z1/length_x1), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    'Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    'Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    'Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    'Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type = ""float"";"
    'BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    'BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    'Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    'Offset"
'    Print #1, "assign OF = " & CStr(valOF) & ", var_type = ""float"";"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type = ""float"";"

    Print #1, "assign theta = 90-2*alpha, var_type = ""float"";"

    Print #1, "assign TempX = W*tan(theta), var_type = ""float"";"
    Print #1, "assign X = sqrt(TempX**2+W**2), var_type = ""float"";"

    Print #1, "assign OriX1 = X*cos(alpha) , var_type = ""float"";"
    Print #1, "assign OriY1 = -X*sin(alpha) , var_type = ""float"";"
    Print #1, "assign OriX2 = -X*cos(alpha) , var_type = ""float"";"
    Print #1, "assign OriY2 = X*sin(alpha) , var_type = ""float"";"

    Print #1, "assign Len = OriY2*2 , var_type = ""float"";"

    Print #1, "assign p11 = OriX1+H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p12 = OriY1+H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p21 = OriX1-H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p22 = OriY1-H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p31 = OriX2-H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p32 = OriY2-H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p41 = OriX2+H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p42 = OriY2+H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p51 = OriX2 , var_type = ""float"";"
    Print #1, "assign p52 = OriY2+Len , var_type = ""float"";"

    Print #1, "assign p61 = OriX1 , var_type = ""float"";"
    Print #1, "assign p62 = OriY1-Len , var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, GPT,"
    Print #1, "vert2 = p61, -p62, GPT,"
    Print #1, "vert3 = p21, -p22, GPT,"
    Print #1, "vert4 = p31, -p32, GPT,"
    Print #1, "vert5 = p51, -p52, GPT,"
    Print #1, "vert6 = p41, -p42, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_XBracing_Type03(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Point at left of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_z = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign length_x1 = (end2_x - end3_x), var_type = ""float"";"
    Print #1, "assign length_z1 = (end3_y - end2_y), var_type = ""float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type = ""float"";"
    'assign alpha1 = atand(length_z1/length_x1), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    'Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    'Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    'Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    'Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type = ""float"";"
    'BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    'BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    'Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    'Offset"
'    Print #1, "assign OF = " & CStr(valOF) & ", var_type = ""float"";"
    '!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type = ""float"";"

    Print #1, "assign theta = 90-2*alpha, var_type = ""float"";"

    Print #1, "assign TempX = W*tan(theta), var_type = ""float"";"
    Print #1, "assign X = sqrt(TempX**2+W**2), var_type = ""float"";"

    Print #1, "assign OriX1 = X*cos(alpha) , var_type = ""float"";"
    Print #1, "assign OriY1 = -X*sin(alpha) , var_type = ""float"";"
    Print #1, "assign OriX2 = -X*cos(alpha) , var_type = ""float"";"
    Print #1, "assign OriY2 = X*sin(alpha) , var_type = ""float"";"

    Print #1, "assign Len = OriY2*2 , var_type = ""float"";"

    Print #1, "assign p11 = OriX1+H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p12 = OriY1+H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p21 = OriX1-H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p22 = OriY1-H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p31 = OriX2-H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p32 = OriY2-H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p41 = OriX2+H*cos(alpha) , var_type = ""float"";"
    Print #1, "assign p42 = OriY2+H*sin(alpha) , var_type = ""float"";"

    Print #1, "assign p51 = OriX2 , var_type = ""float"";"
    Print #1, "assign p52 = OriY2+Len , var_type = ""float"";"

    Print #1, "assign p61 = OriX1 , var_type = ""float"";"
    Print #1, "assign p62 = OriY1-Len , var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, 0,"
    Print #1, "vert2 = p61, p62, 0,"
    Print #1, "vert3 = p21, p22, 0,"
    Print #1, "vert4 = p31, p32, 0,"
    Print #1, "vert5 = p51, p52, 0,"
    Print #1, "vert6 = p41, p42, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_KBracing_Bottom(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Left Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Right Point"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign length_x1 = (end2_x - end3_x), var_type = ""float"";"
    Print #1, "assign length_y1 = (end2_y - end3_y), var_type = ""float"";"

    Print #1, "assign alpha = -1 * atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha1 = atand(length_y1 / length_x1), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BeHD = BeW / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign TempX1 = (BeHD / Sin(alpha1)) + (0.015 / Sin(alpha1)) + (W / Tan(alpha1)), var_type = ""float"";"
    Print #1, "assign H1 = TempX1 + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al1 = H1 * Sin(alpha1), var_type = ""float"";"
    Print #1, "assign Hcos_al1 = H1 * Cos(alpha1), var_type = ""float"";"

    Print #1, "assign Wsin_al1 = W * Sin(alpha1), var_type = ""float"";"
    Print #1, "assign Wcos_al1 = W * Cos(alpha1), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p41 = Hcos_al1 - Wsin_al1, var_type = ""float"";"
    Print #1, "assign p42 = Hsin_al1 + Wcos_al1, var_type = ""float"";"

    Print #1, "assign p51 = Hcos_al1 + Wsin_al1, var_type = ""float"";"
    Print #1, "assign p52 = Hsin_al1 - Wcos_al1, var_type = ""float"";"

    Print #1, "assign p61 = Hcos_al1 + Wsin_al1, var_type = ""float"";"
    Print #1, "assign p62 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, 0,"
    Print #1, "vert2 = p21, p22, 0,"
    Print #1, "vert3 = p31, p32, 0,"
    Print #1, "vert4 = -p41, p42, 0,"
    Print #1, "vert5 = -p51, p52, 0,"
    Print #1, "vert6 = -p61, p62, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
   
End Sub

Public Sub HB_KBracing_Left(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Top Point"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Bottom Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end3_x - end2_x), var_type = ""float"";"
    Print #1, "assign length_y = (end3_y - end2_y), var_type = ""float"";"

    Print #1, "assign alpha = 90-atand(length_y/length_x), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BeHD = BeW/2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hsin_al+Wcos_al, var_type = ""float"";"
    Print #1, "assign p12 = Hcos_al-Wsin_al, var_type = ""float"";"

    Print #1, "assign p21 = Hsin_al-Wcos_al, var_type = ""float"";"
    Print #1, "assign p22 = Hcos_al+Wsin_al, var_type = ""float"";"

    Print #1, "assign p31 = BewtH, var_type = ""float"";"
    Print #1, "assign p32 = Hcos_al+Wsin_al, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p31, p32, " & CStr(valGPT) & ", "
    Print #1, "vert2 = p21, p22, " & CStr(valGPT) & ", "
    Print #1, "vert3 = p11, p12, " & CStr(valGPT) & ", "
    Print #1, "vert4 = p11, -p12, " & CStr(valGPT) & ", "
    Print #1, "vert5 = p21, -p22, " & CStr(valGPT) & ", "
    Print #1, "vert6 = p31, -p32, " & CStr(valGPT) & ", "
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_KBracing_Right(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Top Point"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Bottom Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha = 90 - atand(length_y / length_x), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BeHD = BeW / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p12 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p11 = BewtH, var_type = ""float"";"

    Print #1, "assign p22 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p21 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign p32 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p31 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, 0,"
    Print #1, "vert2 = -p21, p22, 0,"
    Print #1, "vert3 = -p31, p32, 0,"
    Print #1, "vert4 = -p31, -p32, 0,"
    Print #1, "vert5 = -p21, -p22, 0,"
    Print #1, "vert6 = -p11, -p12, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_KBracing_Top(valPathName As String, _
    valBeW As Single, valBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Left Point"";"
    Print #1, "assign end3_x=%%point_x, var_type=""float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Pick Right Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign length_x1 = (end2_x - end3_x), var_type = ""float"";"
    Print #1, "assign length_y1 = (end2_y - end3_y), var_type = ""float"";"

    Print #1, "assign alpha = -1*atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha1 = atand(length_y1/length_x1), var_type = ""float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BeHD = BeW/2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign TempX1 = (BeHD/sin(alpha1))+(0.015/sin(alpha1))+(W/tan(alpha1)), var_type = ""float"";"
    Print #1, "assign H1 = TempX1+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign Hsin_al1 = H1*sin(alpha1), var_type = ""float"";"
    Print #1, "assign Hcos_al1 = H1*cos(alpha1), var_type = ""float"";"

    Print #1, "assign Wsin_al1 = W*sin(alpha1), var_type = ""float"";"
    Print #1, "assign Wcos_al1 = W*cos(alpha1), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BewtH, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign p41 = Hcos_al1-Wsin_al1, var_type = ""float"";"
    Print #1, "assign p42 = Hsin_al1+Wcos_al1, var_type = ""float"";"

    Print #1, "assign p51 = Hcos_al1+Wsin_al1, var_type = ""float"";"
    Print #1, "assign p52 = Hsin_al1-Wcos_al1, var_type = ""float"";"

    Print #1, "assign p61 = Hcos_al1+Wsin_al1, var_type = ""float"";"
    Print #1, "assign p62 = BewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, " & CStr(valGPT) & ", "
    Print #1, "vert2 = p21, -p22, " & CStr(valGPT) & ", "
    Print #1, "vert3 = p31, -p32, " & CStr(valGPT) & ", "
    Print #1, "vert4 = -p41, -p42, " & CStr(valGPT) & ", "
    Print #1, "vert5 = -p51, -p52, " & CStr(valGPT) & ", "
    Print #1, "vert6 = -p61, -p62, " & CStr(valGPT) & ", "
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_BtoB_LeftBottom_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valSBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' SubBeam Web Thickness"
    Print #1, "assign SBewt = " & CStr(valSBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"
    Print #1, "assign SBewtH = SBewt / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign X = 1/(tan(alpha)**2), var_type = ""float"";"
    Print #1, "assign XX = sqrt(X + 1), var_type = ""float"";"
    Print #1, "assign TempX = W * XX, var_type = ""float"";"

    Print #1, "assign p41 = BewtH + TempX, var_type = ""float"";"
    Print #1, "assign p42 = SBewtH, var_type = ""float"";"

    Print #1, "assign p51 = BewtH, var_type = ""float"";"
    Print #1, "assign p52 = SBewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, p12, GPT,"
    Print #1, "vert2 = p21, p22, GPT,"
    Print #1, "vert3 = p31, p32, GPT,"
    Print #1, "vert4 = p41, p42, GPT,"
    Print #1, "vert5 = p51, p52, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_BtoB_LeftTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valSBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' SubBeam Web Thickness"
    Print #1, "assign SBewt = " & CStr(valSBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"
    Print #1, "assign SBewtH = SBewt/2 ,var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign X = 1/(tan(alpha)**2), var_type = ""float"";"
    Print #1, "assign XX = sqrt(X+1), var_type = ""float"";"
    Print #1, "assign TempX = w*XX, var_type = ""float"";"

    Print #1, "assign p41 = BewtH+TempX, var_type = ""float"";"
    Print #1, "assign p42 = SBewtH, var_type = ""float"";"

    Print #1, "assign p51 = BewtH, var_type = ""float"";"
    Print #1, "assign p52 = SBewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = p11, -p12, 0,"
    Print #1, "vert2 = p21, -p22, 0,"
    Print #1, "vert3 = p31, -p32, 0,"
    Print #1, "vert4 = p41, -P42, 0,"
    Print #1, "vert5 = p51, -P52, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_BtoB_RightBottom_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valSBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' SubBeam Web Thickness"
    Print #1, "assign SBewt = " & CStr(valSBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD / 2, var_type = ""float"";"
    Print #1, "assign BewtH = Bewt / 2, var_type = ""float"";"
    Print #1, "assign BeWH = BeW / 2, var_type = ""float"";"
    Print #1, "assign SBewtH = SBewt / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeWH / Cos(alpha)) + (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al + Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign X = 1/(tan(alpha)**2), var_type = ""float"";"
    Print #1, "assign XX = sqrt(X + 1), var_type = ""float"";"
    Print #1, "assign TempX = W * XX, var_type = ""float"";"

    Print #1, "assign p41 = BewtH + TempX, var_type = ""float"";"
    Print #1, "assign p42 = SBewtH, var_type = ""float"";"

    Print #1, "assign p51 = BewtH, var_type = ""float"";"
    Print #1, "assign p52 = SBewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, p12, 0,"
    Print #1, "vert2 = -p21, p22, 0,"
    Print #1, "vert3 = -p31, p32, 0,"
    Print #1, "vert4 = -p41, P42, 0,"
    Print #1, "vert5 = -p51, p52, 0,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub

Public Sub HB_BtoB_RightTop_Center(valPathName As String, _
    valBeW As Single, valBewt As Single, valSBewt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single)
    
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""float"";"

    Print #1, "assign length_x = (end2_x - end1_x), var_type = ""float"";"
    Print #1, "assign length_y = (end2_y - end1_y), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_y/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Beam Width"
    Print #1, "assign BeW = " & CStr(valBeW) & ", var_type = ""float"";"
    ' Beam Web Thickness"
    Print #1, "assign Bewt = " & CStr(valBewt) & ", var_type = ""float"";"
    ' SubBeam Web Thickness"
    Print #1, "assign SBewt = " & CStr(valSBewt) & ", var_type = ""float"";"
    ' Bracing Width"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type = ""float"";"
    ' Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type = ""float"";"
    ' BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"
    ' BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    ' Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type = ""float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BewtH = Bewt/2 ,var_type = ""float"";"
    Print #1, "assign BeWH = BeW/2, var_type = ""float"";"
    Print #1, "assign SBewtH = SBewt/2 ,var_type = ""float"";"

    Print #1, "assign TempX = (BeWH/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = BewtH, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign X = 1/(tan(alpha)**2), var_type = ""float"";"
    Print #1, "assign XX = sqrt(X+1), var_type = ""float"";"
    Print #1, "assign TempX = w*XX, var_type = ""float"";"

    Print #1, "assign p41 = BewtH+TempX, var_type = ""float"";"
    Print #1, "assign p42 = SBewtH, var_type = ""float"";"

    Print #1, "assign p51 = BewtH, var_type = ""float"";"
    Print #1, "assign p52 = SBewtH, var_type = ""float"";"

    Print #1, "plc_area"
    Print #1, "vert1 = -p11, -p12, GPT,"
    Print #1, "vert2 = -p21, -p22, GPT,"
    Print #1, "vert3 = -p31, -p32, GPT,"
    Print #1, "vert4 = -p41, -P42, GPT,"
    Print #1, "vert5 = -p51, -P52, GPT,"
    Print #1, "class = " & gstr_HBClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""GP_" & CStr(valGPT) & """" & ", "
    Print #1, "thickness = GPT;"
Close #1
End Sub
Public Sub Hor_HTB_Nut_M01(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (BeWH/Sin(alpha)) + (0.015/Sin(alpha)) + (W/Tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M02(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If

               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M03(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M04(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M05_1(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, _
                                                                   valNutHei As Single, valHTBNum As Integer, valHTBSpace As Single, valGPThk As Single, _
                                                                   valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = W+SP3, var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M05_2(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, _
                                                                   valNutHei As Single, valHTBNum As Integer, valHTBSpace As Single, valGPThk As Single, _
                                                                   valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = W+SP3, var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
'               Print #1, "plc_area"
'               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "class = " & gstr_HBClass & ", " & _
'                         "grade = """ & gstr_Grade & """, " & _
'                         "material = """ & gstr_Material & """, " & _
'                         "name = ""HTB_" & valBoltName & """, "
'               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
'               Print #1, "plc_area"
'               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "class = " & gstr_HBClass & ", " & _
'                         "grade = """ & gstr_Grade & """, " & _
'                         "material = """ & gstr_Material & """, " & _
'                         "name = ""HTB_" & valBoltName & """, "
'               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
'                              Print #1, "plc_area"
'                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "class = " & gstr_HBClass & ", " & _
'                                        "grade = """ & gstr_Grade & """, " & _
'                                        "material = """ & gstr_Material & """, " & _
'                                        "name = ""HTB_" & valBoltName & """, "
'                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
'                              Print #1, "plc_area"
'                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "class = " & gstr_HBClass & ", " & _
'                                        "grade = """ & gstr_Grade & """, " & _
'                                        "material = """ & gstr_Material & """, " & _
'                                        "name = ""HTB_" & valBoltName & """, "
'                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M05_3(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, _
                                                                   valNutHei As Single, valHTBNum As Integer, valHTBSpace As Single, valGPThk As Single, _
                                                                   valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = W+SP3, var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
'               Print #1, "plc_area"
'               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "class = " & gstr_HBClass & ", " & _
'                         "grade = """ & gstr_Grade & """, " & _
'                         "material = """ & gstr_Material & """, " & _
'                         "name = ""HTB_" & valBoltName & """, "
'               Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'               Print #1, "plc_area"
'               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'               Print #1, "class = " & gstr_HBClass & ", " & _
'                         "grade = """ & gstr_Grade & """, " & _
'                         "material = """ & gstr_Material & """, " & _
'                         "name = ""HTB_" & valBoltName & """, "
'               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
'                              Print #1, "plc_area"
'                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "class = " & gstr_HBClass & ", " & _
'                                        "grade = """ & gstr_Grade & """, " & _
'                                        "material = """ & gstr_Material & """, " & _
'                                        "name = ""HTB_" & valBoltName & """, "
'                              Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                              Print #1, "plc_area"
'                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
'                              Print #1, "class = " & gstr_HBClass & ", " & _
'                                        "grade = """ & gstr_Grade & """, " & _
'                                        "material = """ & gstr_Material & """, " & _
'                                        "name = ""HTB_" & valBoltName & """, "
'                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_a(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, _
                                                                  valNutHei As Single, ByVal valHTBNum As Integer, _
                                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH/Sin(alpha)) + (0.015/Sin(alpha)) + (W/Tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_b1(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, _
                                                                     valNutHei As Single, ByVal valHTBNum As Integer, valHTBSpace As Single, valGPThk As Single, _
                                                                     valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH/Sin(alpha)) + (0.015/Sin(alpha)) + (W/Tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
             
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_b2(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                     valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                                     valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX1+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha1 + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha1 - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX1+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX1+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha1 + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha1 - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX1+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If

               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_c(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If

               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
                              
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_d1(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                     valNutDia As Single, valNutHei As Single, ByVal valHTBNum As Integer, _
                                                                     valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               Print #1, "plc_area"
               Print #1, "vert1 = -HTB_X1, -HTB_Y1," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert2 = -HTB_X2, -HTB_Y2," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert3 = -HTB_X3, -HTB_Y3," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert4 = -HTB_X4, -HTB_Y4," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert5 = -HTB_X5, -HTB_Y5," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "vert6 = -HTB_X6, -HTB_Y6," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = -HTB_X11, -HTB_Y11," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert2 = -HTB_X21, -HTB_Y21," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert3 = -HTB_X31, -HTB_Y31," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert4 = -HTB_X41, -HTB_Y41," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert5 = -HTB_X51, -HTB_Y51," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "vert6 = -HTB_X61, -HTB_Y61," & CStr(Format(valGPThk + valNutHei, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub
Public Sub Hor_HTB_Nut_M06_d2(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                     valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                                     valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
'    Print #1, "assign TempX = (BeWH / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign seta1 = 60, var_type = ""float"";"
    Print #1, "assign seta2 = 120, var_type = ""float"";"
    Print #1, "assign seta3 = 180, var_type = ""float"";"
    Print #1, "assign seta4 = 240, var_type = ""float"";"
    Print #1, "assign seta5 = 300, var_type = ""float"";"
    Print #1, "assign seta6 = 360, var_type = ""float"";"
    
    If valBoltType = "II" Then valHTBNum = valHTBNum / 2
    
    For i = 1 To valHTBNum
               If i = 1 Then
                              
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX1+SP1, var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha1 + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha1 - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX1+SP1, var_type = ""float"";"
                              End If
                              
               Else
                              If valBoltType = "II" Then
                                             Print #1, "assign HH = TempX1+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                                             Print #1, "assign ceta1 = alpha1 + atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign ceta2 = alpha1 - atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
                                             Print #1, "assign AA = (HH)**2, var_type = ""float"";"
                                             Print #1, "assign BB = " & CStr(valGage) & "**2, var_type = ""float"";"
                                             Print #1, "assign HH = sqrt(AA+BB), var_type = ""float"";"
                              Else
                                             Print #1, "assign HH = TempX1+SP1+" & CStr(Format(valHTBSpace * (i - 1), "0.000")) & ", var_type = ""float"";"
                              End If
                              
               End If
               
               If valBoltType = "II" Then
                              Print #1, "assign Hx = HH*cos(ceta1), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(ceta1), var_type = ""float"";"
                              Print #1, "assign Hx1 = HH*cos(ceta2), var_type = ""float"";"
                              Print #1, "assign Hy1 = HH*sin(ceta2), var_type = ""float"";"
               Else
                              Print #1, "assign Hx = HH*cos(alpha), var_type = ""float"";"
                              Print #1, "assign Hy = HH*sin(alpha), var_type = ""float"";"
                              
               End If
               
               Print #1, "assign HTB_X1= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
               Print #1, "assign HTB_Y1= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
               Print #1, "assign HTB_X2= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
               Print #1, "assign HTB_Y2= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
               Print #1, "assign HTB_X3= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
               Print #1, "assign HTB_Y3= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
               Print #1, "assign HTB_X4= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
               Print #1, "assign HTB_Y4= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
               Print #1, "assign HTB_X5= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
               Print #1, "assign HTB_Y5= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
               Print #1, "assign HTB_X6= Hx+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
               Print #1, "assign HTB_Y6= Hy+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               
               If valBoltType = "II" Then
                              Print #1, "assign HTB_X11= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_Y11= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta1), var_type = ""float"";"
                              Print #1, "assign HTB_X21= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_Y21= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta2), var_type = ""float"";"
                              Print #1, "assign HTB_X31= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_Y31= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta3), var_type = ""float"";"
                              Print #1, "assign HTB_X41= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_Y41= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta4), var_type = ""float"";"
                              Print #1, "assign HTB_X51= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_Y51= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta5), var_type = ""float"";"
                              Print #1, "assign HTB_X61= Hx1+" & CStr(Format(valNutDia / 2, "0.000")) & "*sin(seta6), var_type = ""float"";"
                              Print #1, "assign HTB_Y61= Hy1+" & CStr(Format(valNutDia / 2, "0.000")) & "*cos(seta6), var_type = ""float"";"
               End If
               
               Print #1, "plc_area"
               Print #1, "vert1 = HTB_X1, -HTB_Y1," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert2 = HTB_X2, -HTB_Y2," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert3 = HTB_X3, -HTB_Y3," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert4 = HTB_X4, -HTB_Y4," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert5 = HTB_X5, -HTB_Y5," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "vert6 = HTB_X6, -HTB_Y6," & CStr(Format(valGPThk, "0.000")) & ","
               Print #1, "class = " & gstr_HBClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""HTB_" & valBoltName & """, "
               Print #1, "thickness = " & CStr(valNutHei) & ";"
               
               If valBoltType = "II" Then
                              Print #1, "plc_area"
                              Print #1, "vert1 = HTB_X11, -HTB_Y11," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert2 = HTB_X21, -HTB_Y21," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert3 = HTB_X31, -HTB_Y31," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert4 = HTB_X41, -HTB_Y41," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert5 = HTB_X51, -HTB_Y51," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "vert6 = HTB_X61, -HTB_Y61," & CStr(Format(valGPThk, "0.000")) & ","
                              Print #1, "class = " & gstr_HBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""HTB_" & valBoltName & """, "
                              Print #1, "thickness = " & CStr(valNutHei) & ";"
               End If
    Next i
    
Close #1
End Sub

Public Sub HBData_Call(ByVal valJobName As String, ByVal valCode As String, ByVal valFormCode As String, _
                                                valBeam As String, valSubBeam As String, valBracing As String)
Dim xSQL As String
Dim reData As ADODB.Recordset

If valBeam <> "N/A" Then
    xSQL = "select * from code_" & valFormCode & " where member_name = '" & valBeam & "'"
    
    Set retData = adoConnection.Execute(xSQL)
        gsin_Bedepth = gf_StringtoSingle(retData!D)
        gsin_Bewidth = gf_StringtoSingle(retData!Bf)
        gsin_Bewt = gf_StringtoSingle(retData!Tw)
        gsin_Beft = gf_StringtoSingle(retData!Tf)
    retData.Close
    
    Set retData = Nothing
Else
    gsin_Bedepth = 0
    gsin_Bewidth = 0
    gsin_Bewt = 0
    gsin_Beft = 0
End If

If valSubBeam <> "N/A" Then
    xSQL = "select * from code_" & valFormCode & " where member_name = '" & valSubBeam & "'"
    
    Set retData = adoConnection.Execute(xSQL)
        gsin_SubBedepth = gf_StringtoSingle(retData!D)
        gsin_SubBewidth = gf_StringtoSingle(retData!Bf)
        gsin_SubBewt = gf_StringtoSingle(retData!Tw)
        gsin_SubBeft = gf_StringtoSingle(retData!Tf)
    retData.Close
    
    Set retData = Nothing
Else
    gsin_SubBedepth = 0
    gsin_SubBewidth = 0
    gsin_SubBewt = 0
    gsin_SubBeft = 0
End If

xSQL = "select * from code_" & valCode & " where member_name = '" & valBracing & "'"

Set retData = adoConnection.Execute(xSQL)
    gsin_Bdepth = gf_StringtoSingle(retData!D)
    gsin_Bwidth = gf_StringtoSingle(retData!Bf)
    gsin_Bwt = gf_StringtoSingle(retData!Tw)
    gsin_Bft = gf_StringtoSingle(retData!Tf)
retData.Close

Set retData = Nothing

xSQL = "select * from HB_Connection where member_name = '" & valBracing & "'"
xSQL = xSQL & "and code = '" & valCode & "' "
xSQL = xSQL & "and job = '" & valJobName & "'"

Set retData = adoConnection1.Execute(xSQL)
'    gsin_SP1 = retData!Space1
'    gsin_SP2 = retData!Space2
''    gsin_SP3 = retData!Space3
'    gsin_HTB_SNum = retData!HTB_SNum
'    gsin_HTB_Space = retData!HTB_Space
'    gsin_GPThk = retData!Plate_Thk
'    gstr_Unit = retData!unit
    
    gsin_SP1 = retData!Space1
    gsin_SP2 = retData!Space2
'    gsin_SP3 = retData!Space3
    gstr_BoltType = retData!Type
    gstr_BoltName = retData!HTB_Name
    gsin_HTB_Num = retData!HTB_Num
    gsin_HTB_SNum = retData!HTB_SNum
    gsin_HTB_Space = retData!HTB_Space
    gsin_GPThk = retData!Plate_Thk
    gsin_Gage = retData!Gage
    gstr_Unit = retData!unit
retData.Close

Set retData = Nothing

xSQL = "select dia,nutdia,nuthei,unit from BoltNut"
xSQL = xSQL & " where Name = '" & gstr_BoltName & "' and unit = '" & gstr_Unit & "'"
Set reData1 = adoConnection.Execute(xSQL)
    gsin_BoltDia = reData1!dia
    gsin_NutDia = reData1!nutdia
    gsin_NutHei = reData1!nuthei
    gstr_NutUnit = reData1!unit
reData1.Close
Set reData1 = Nothing


xSQL = "select Grade,Material,HB_Class from Plate_General "
xSQL = xSQL & "where job = '" & valJobName & "'"

Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    gstr_Grade = reData1!grade
    gstr_Material = reData1!material
    gstr_HBClass = reData1!hb_class
Else
    gstr_Grade = "A36"
    gstr_Material = "Steel"
    gstr_HBClass = "3"
End If
reData1.Close
Set reData1 = Nothing

End Sub


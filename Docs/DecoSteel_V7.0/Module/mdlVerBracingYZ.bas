Attribute VB_Name = "mdlVerBracingYZ"

Public Sub VB_PML_YZ(valPathName As String, ByVal valJobName As String, ByVal valCode As String, _
    ByVal valFormCode As String, valVBGP_Flag As String, valColShape As String, valType As String, valUnit As String, _
    valColumn As String, valBeam As String, valBracing As String, valBracing_Type As String, _
    valBracitg_Dir As String, valTempSpace3 As Single, valNutFlag As Integer)


Dim TempCdepth As Single, TempCwidth As Single, TempCft As Single, TempCwt As Single
Dim TempBedepth As Single, TempBdepth As Single, TempBwt As Single
Dim TempSP1 As Single, TempSP2 As Single, TempSP3 As Single
Dim TempHTB_Space As Single, TempGPThk As Single, TempOF As Single
Dim Br_Flag As String
Dim TempBoltName As String, TempNutDia As Single, TempNutHei As Single, TempGage As Single

    Call VBData_Call(valJobName, valCode, valFormCode, valColumn, valBeam, valBracing)
    Br_Flag = valBracing_Type & valBracitg_Dir

    TempCdepth = convert(gsin_Cdepth, "mm", valUnit)
    TempCwidth = convert(gsin_Cwidth, "mm", valUnit)
    TempCft = convert(gsin_Cft, "mm", valUnit)
    TempCwt = convert(gsin_Cwt, "mm", valUnit)
    TempBedepth = convert(gsin_Bedepth, "mm", valUnit)
    TempBdepth = convert(gsin_Bdepth, "mm", valUnit)
    TempBwt = convert(gsin_Bwt, "mm", valUnit)

    TempSP1 = convert(gsin_SP1, gstr_Unit, valUnit)
    TempSP2 = convert(gsin_SP2, gstr_Unit, valUnit)
    TempSP3 = valTempSpace3
    TempHTB_Space = convert(gsin_HTB_Space, gstr_Unit, valUnit)
    TempGPThk = convert(gsin_GPThk, gstr_Unit, valUnit)
    TempGage = convert(gsin_Gage, gstr_Unit, valUnit)
    
    TempBoltName = gstr_BoltName
    TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
    TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)

    Select Case valVBGP_Flag
        Case "Module01"
            If valColShape = "Strong Axis" Then
                Select Case valType
                    Case "Type-01"
                              Call VB_STRONG_LeftBottJoint_YZ(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                              
                    Case "Type-02"
                              Call VB_STRONG_LeftBott_YZ(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    
                    Case "Type-03"
                              Call VB_STRONG_LeftBott_YZ_Type03(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-04"

                              Call VB_STRONG_LeftBott_YZ_Type04(valPathName, TempBedepth, TempBdepth, _
                                                   TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    
                              
                End Select
            Else
                Select Case valType
                    Case "Type-01"
                              Call VB_WEAK_LeftBottJoint_YZ(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                              
                    Case "Type-02"
                              Call VB_WEAK_LeftBott_YZ(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                              Call VB_WEAK_LeftBott_YZ_Type03(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            End If
               If valNutFlag = 1 Then
                              Call Ver_HTB_Nut_M01_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
               End If
        Case "Module02"
            If valColShape = "Strong Axis" Then
                Select Case valType
                    Case "Type-01"
                              Call VB_STRONG_RightBottJoint_YZ(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                              Call VB_STRONG_RightBott_YZ(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                              Call VB_STRONG_RightBott_YZ_Type03(valPathName, TempCdepth, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-04"
                              Call VB_STRONG_RightBott_YZ_Type04(valPathName, TempBedepth, TempBdepth, _
                                                               TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    
                End Select
            Else
                Select Case valType
                    Case "Type-01"
                              Call VB_WEAK_RightBottJoint_YZ(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                              Call VB_WEAK_RightBott_YZ(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                              Call VB_WEAK_RightBott_YZ_Type03(valPathName, TempCwidth, TempCwt, TempBdepth, _
                                             TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            End If
               If valNutFlag = 1 Then
                              Call Ver_HTB_Nut_M02_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
               End If
        Case "Module03"
            If valColShape = "Strong Axis" Then
                Select Case valType
                    Case "Type-01"
                              Call VB_STRONG_LeftTop_YZ_Type01(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                              TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                        Call VB_STRONG_LeftTop_YZ_Type02(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                        Call VB_STRONG_LeftTop_YZ_Type03(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-04"
                        Call VB_STRONG_LeftTop_YZ_Type04(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-05"
                        Call VB_STRONG_LeftTop_YZ_Type05(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-06"
                        Call VB_STRONG_LeftTop_YZ_Type06(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            Else
                Select Case valType
                    Case "Type-01"
                        Call VB_WEAK_LeftTop_YZ_Type01(valPathName, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                        Call VB_WEAK_LeftTop_YZ_Type02(valPathName, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                        Call VB_WEAK_LeftTop_YZ_Type03(valPathName, TempCwidth, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            End If
               If valNutFlag = 1 Then
                              Call Ver_HTB_Nut_M03_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
               End If
        Case "Module04"
            If valColShape = "Strong Axis" Then
                Select Case valType
                    Case "Type-01"
                         Call VB_STRONG_RightTop_YZ_Type01(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                         Call VB_STRONG_RightTop_YZ_Type02(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                         Call VB_STRONG_RightTop_YZ_Type03(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-04"
                         Call VB_STRONG_RightTop_YZ_Type04(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-05"
                         Call VB_STRONG_RightTop_YZ_Type05(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-06"
                         Call VB_STRONG_RightTop_YZ_Type06(valPathName, TempCdepth, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            Else
                Select Case valType
                    Case "Type-01"
                         Call VB_WEAK_RightTop_YZ_Type01(valPathName, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-02"
                         Call VB_WEAK_RightTop_YZ_Type02(valPathName, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    Case "Type-03"
                         Call VB_WEAK_RightTop_YZ_Type03(valPathName, TempCwidth, TempCwt, TempBedepth, TempBdepth, _
                            TempSP1, TempSP2, gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                End Select
            End If
               If valNutFlag = 1 Then
                              Call Ver_HTB_Nut_M04_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
               End If
        Case "Module05"
            Select Case valType
                Case "Type-01"
                    Call VB_Offset_LeftTop_YZ_Type01(valPathName, _
                        TempBedepth, TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    If valNutFlag = 1 Then
                        Call Ver_HTB_Nut_M03_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                    End If
                Case "Type-02"
                    Call VB_Offset_LeftTop_YZ_Type02(valPathName, _
                        TempBedepth, TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    If valNutFlag = 1 Then
                        Call Ver_HTB_Nut_M05_YZ_2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                    End If
            End Select
               
        Case "Module06"
            Select Case valType
                Case "Type-01"
                    Call VB_Offset_RightTop_YZ_Type01(valPathName, _
                         TempBedepth, TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    If valNutFlag = 1 Then
                        Call Ver_HTB_Nut_M04_YZ(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                    End If
                Case "Type-02"
                    Call VB_Offset_RightTop_YZ_Type02(valPathName, _
                        TempBedepth, TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                    If valNutFlag = 1 Then
                        Call Ver_HTB_Nut_M06_YZ_2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                    End If
            End Select
               
        Case "Module07"
            Select Case valType
                Case "Type-01"
                              Call VB_XBracing_YZ_Type01(valPathName, _
                                             TempBdepth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                              If valNutFlag = 1 Then
                                             Call Ver_HTB_Nut_M07_YZ_1(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                              End If
                Case "Type-02"
                              Call VB_XBracing_YZ_Type02(valPathName, _
                                             TempBdepth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                              If valNutFlag = 1 Then
                                             Call Ver_HTB_Nut_M07_YZ_2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                              End If
                Case "Type-03"
                              Call VB_XBracing_YZ_Type03(valPathName, _
                                             TempBdepth, TempSP1, TempSP2, TempSP3, _
                                             gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                              If valNutFlag = 1 Then
                                             Call Ver_HTB_Nut_M07_YZ_3(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                                            TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                              End If
            End Select
        Case "Module08"
            Select Case valType
                Case "Type-01"
                    Call VB_KBracing_YZ_Type01(valPathName, _
                        TempBedepth, TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
                Case "Type-02"
                    Call VB_KBracing_YZ_Type02(valPathName, _
                         TempBdepth, TempSP1, TempSP2, _
                        gsin_HTB_SNum, TempHTB_Space, TempGPThk, Br_Flag)
            End Select
               If valNutFlag = 1 Then
                              Call Ver_HTB_Nut_M08_YZ_1(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
                              Call Ver_HTB_Nut_M08_YZ_2(valPathName, gstr_BoltType, TempBoltName, TempNutDia, _
                                             TempNutHei, gsin_HTB_Num, TempHTB_Space, TempGPThk, TempGage, TempBwt, Br_Flag)
               End If
    End Select
End Sub

Public Sub VB_STRONG_RightBottJoint_YZ(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(30), var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
    
Close #1
End Sub

Public Sub VB_STRONG_RightBott_YZ(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type=""Float"";"

    ' !assign p11 = 0+CHD, var_type = "float";
    ' !assign p12 = 0, var_type = "float";
    ' !assign p21 = 0+CHD, var_type = "float";
    ' !assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
    ' !assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
    ' !assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
    ' !assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
    ' !assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
    ' !assign p51 = Len*cos(alpha)+CHD, var_type = "float";
    ' !assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = -1*Len*cos(alpha), var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = -1*Len*cos(alpha), var_type=""Float"";"

'    Print #1, "plc_area"
'    Print #1, "vert1 = 0, p11, p12,"
'    Print #1, "vert2 = 0, p21, p22,"
'    Print #1, "vert3 = 0, p31, p32,"
'    Print #1, "vert4 = 0, p41, p42,"
'    Print #1, "vert5 = 0, p51, p52,"
'    Print #1, "class = " & gstr_VBClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""GP_" & CStr(valGPT) & """" & ", "
'    Print #1, "thickness = GPT;"
'
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub
Public Sub VB_STRONG_RightBott_YZ_Type03(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = -end1_z, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = -end1_z, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub
Public Sub VB_STRONG_RightBott_YZ_Type04(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String, _
    Optional valCD As Single)

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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z / length_x), var_type = ""float"";"
    Print #1, "assign alpha =  alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
'    If valCD <> 0 Then
'      Print #1, "assign CHD = " & CStr(valCD / 2) & ", var_type = ""float"";"
'    End If
        
    ' Beam Depth"
    Print #1, "assign Bed = " & CStr(valBed) & ", var_type = ""float"";"
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
    Print #1, "assign W = " & Format(CStr(valBd / 2), "0.000") & ",var_type = ""Float"";"

    Print #1, "assign TempX = (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = 0, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

'    If valCD <> 0 Then
'      Print #1, "assign p41 = CHD, var_type = ""float"";"
'    Else
      Print #1, "assign alpha2 = (90 - alpha), var_type = ""float"";"
      Print #1, "assign Len = p32, var_type = ""float"";"
      Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"
      Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
'    End If
    Print #1, "assign p42 = 0, var_type = ""float"";"
    
    Print #1, "assign TempX = (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
'                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
'                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
'                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
'                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, -p42,"
                              
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
'                              Print #1, "vert1 = 0, p11, -p12,"
'                              Print #1, "vert2 = 0, p21, p22,"
'                              Print #1, "vert3 = 0, p31, p32,"
'                              Print #1, "vert4 = 0, p41, -p42,"
                              
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
'                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
'                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
'                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
'                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1



End Sub

Public Sub VB_STRONG_LeftBottJoint_YZ(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(30), var_type=""Float"";"

' !assign p11 = 0+CHD, var_type = "float";
' !assign p12 = 0, var_type = "float";
' !assign p21 = 0+CHD, var_type = "float";
' !assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
' !assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
' !assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
' !assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
' !assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
' !assign p51 = Len*cos(alpha)+CHD, var_type = "float";
' !assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

'    Print #1, "plc_area"
'    Print #1, "vert1 = 0, -p11, p12,"
'    Print #1, "vert2 = 0, -p21, p22,"
'    Print #1, "vert3 = 0, -p31, p32,"
'    Print #1, "vert4 = 0, -p41, p42,"
'    Print #1, "vert5 = 0, -p51, p52,"
'    Print #1, "class = " & gstr_VBClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""GP_" & CStr(valGPT) & """" & ", "
'    Print #1, "thickness = GPT;"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
    
Close #1
End Sub

Public Sub VB_STRONG_LeftBott_YZ(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type=""Float"";"

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = -1*Len*cos(alpha), var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = -1*Len*cos(alpha), var_type=""Float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub
Public Sub VB_STRONG_LeftBott_YZ_Type03(valPathName As String, _
    valCD As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = (Len*sin(alpha))+CHD, var_type=""Float"";"
    Print #1, "assign p42 = -end1_z, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = -end1_z, var_type=""Float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub
Public Sub VB_STRONG_LeftBott_YZ_Type04(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String, _
    Optional valCD As Single)

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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z / length_x), var_type = ""float"";"
    Print #1, "assign alpha =  -1*alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
'    If valCD <> 0 Then
'      Print #1, "assign CHD = " & CStr(valCD / 2) & ", var_type = ""float"";"
'    End If
        
    ' Beam Depth"
    Print #1, "assign Bed = " & CStr(valBed) & ", var_type = ""float"";"
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
    Print #1, "assign W = " & Format(CStr(valBd / 2), "0.000") & ",var_type = ""Float"";"

    Print #1, "assign TempX = (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = 0, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

'    If valCD <> 0 Then
'      Print #1, "assign p41 = CHD, var_type = ""float"";"
'    Else
      Print #1, "assign alpha2 = (90 - alpha), var_type = ""float"";"
      Print #1, "assign Len = p32, var_type = ""float"";"
      Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"
      Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
'    End If
    Print #1, "assign p42 = 0, var_type = ""float"";"
    
    Print #1, "assign TempX = (0.015 / Cos(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
'                              Print #1, "vert1 = p11, 0, -p12,"
'                              Print #1, "vert2 = p21, 0, p22,"
'                              Print #1, "vert3 = p31, 0, p32,"
'                              Print #1, "vert4 = p41,0, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, -p42,"
                              
'                              Print #1, "vert1 = p11, -" & CStr(valGPT) & ", -p12,"
'                              Print #1, "vert2 = p21, -" & CStr(valGPT) & ", p22,"
'                              Print #1, "vert3 = p31, -" & CStr(valGPT) & ", p32,"
'                              Print #1, "vert4 = p41, -" & CStr(valGPT) & ", -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              
'                              Print #1, "vert1 = p11, -" & CStr(Format(valGPT / 2, "0.000")) & ", -p12,"
'                              Print #1, "vert2 = p21, -" & CStr(Format(valGPT / 2, "0.000")) & ", p22,"
'                              Print #1, "vert3 = p31, -" & CStr(Format(valGPT / 2, "0.000")) & ", p32,"
'                              Print #1, "vert4 = p41, -" & CStr(Format(valGPT / 2, "0.000")) & ", -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1


End Sub
Public Sub VB_STRONG_RightTop_YZ_Type01(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = BeD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeD, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_STRONG_RightTop_YZ_Type02(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"
    Print #1, "assign BeHD = BeD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = BeHD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeHD, var_type=""Float"";"

        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_STRONG_RightTop_YZ_Type03(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CD/(sin(alpha)*2))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_STRONG_RightTop_YZ_Type04(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = BeD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeD, var_type=""Float"";"

        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_STRONG_RightTop_YZ_Type05(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"
    Print #1, "assign BeHD = BeD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = BeHD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeHD, var_type=""Float"";"

        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_STRONG_RightTop_YZ_Type06(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = -1*alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type01(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = BeD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeD, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type02(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"
    Print #1, "assign BeHD = BeD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
    ' assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type=""Float"";"

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = BeHD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeHD, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type03(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (CD/(sin(alpha)*2))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = CHD, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = CHD, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type04(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = BeD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeD, var_type=""Float"";"
 
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type05(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"
    Print #1, "assign BeHD = BeD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = BeHD, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = BeHD, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_STRONG_LeftTop_YZ_Type06(valPathName As String, _
    valCD As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Define start Point"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Define End Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha = alpha1, var_type=""Float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !Column Depth"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type=""Float"";"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign CHD = CD/2, var_type=""Float"";"

    Print #1, "assign TempX = (0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"
'!assign Len = ((W*tan(alpha))+H-(CHD/cos(alpha)))*sin(15), var_type = "float";

'!assign p11 = 0+CHD, var_type = "float";
'!assign p12 = 0, var_type = "float";
'!assign p21 = 0+CHD, var_type = "float";
'!assign p22 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p31 = H*sin(alpha)-W*cos(alpha)+CHD, var_type = "float";
'!assign p32 = W*sin(alpha)+H*cos(alpha), var_type = "float";
'!assign p41 = H*sin(alpha)+W*cos(alpha)+CHD, var_type = "float";
'!assign p42 = H*cos(alpha)-W*sin(alpha), var_type = "float";
'!assign p51 = Len*cos(alpha)+CHD, var_type = "float";
'!assign p52 = 0, var_type = "float";

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign p11 = 0, var_type=""Float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p51 = 0, var_type=""Float"";"
    Print #1, "assign p52 = 0, var_type=""Float"";"

    Print #1, "assign Len = P32-P52, var_type=""Float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type=""Float"";"

    Print #1, "assign p41 = P31-TempXX, var_type=""Float"";"
    Print #1, "assign p42 = 0, var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_XBracing_YZ_Type01(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Left of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign length_x1 = (end2_y - end3_y), var_type=""Float"";"
    Print #1, "assign length_z1 = (end3_z - end2_z), var_type=""Float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type=""Float"";"
'!assign alpha1 = atand(length_z1/length_x1), var_type = "float";

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type=""Float"";"
    Print #1, "assign alpha1 = atand(W/H), var_type=""Float"";"
    Print #1, "assign Len = sqrt(H**2+W**2), var_type=""Float"";"

    Print #1, "assign Len1 = sqrt(0.01**2+W**2), var_type=""Float"";"
    Print #1, "assign alpha2 = atand(W/0.1), var_type=""Float"";"
    Print #1, "assign alpha3 = 90-alpha2+alpha, var_type=""Float"";"

    Print #1, "assign p11 = len*cos(alpha+alpha1), var_type=""Float"";"
    Print #1, "assign p12 = len*sin(alpha+alpha1), var_type=""Float"";"

    Print #1, "assign p21 = len*cos(alpha-alpha1), var_type=""Float"";"
    Print #1, "assign p22 = len*sin(alpha-alpha1), var_type=""Float"";"

    Print #1, "assign p31 = p21-len1*sin(alpha3), var_type=""Float"";"
    Print #1, "assign p32 = p22-len1*cos(alpha3), var_type=""Float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert6 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert7 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert8 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p11, p12,"
                              Print #1, "vert3 = 0, p21, p22,"
                              Print #1, "vert4 = 0, -p21, p22,"
                              Print #1, "vert5 = 0, -p11, p12,"
                              Print #1, "vert6 = 0, -p11, -p12,"
                              Print #1, "vert7 = 0, -p21, -p22,"
                              Print #1, "vert8 = 0, p21, -p22,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert6 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert7 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22,"
                              Print #1, "vert8 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

Public Sub VB_XBracing_YZ_Type03(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Left of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign length_x1 = (end2_y - end3_y), var_type=""Float"";"
    Print #1, "assign length_z1 = (end3_z - end2_z), var_type=""Float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type=""Float"";"
' !assign alpha1 = atand(length_z1/length_x1), var_type = "float";

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !Offset"
'    Print #1, "assign OF = " & CStr(valOF) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type=""Float"";"

    Print #1, "assign theta = 90-2*alpha, var_type=""Float"";"

    Print #1, "assign TempX = W*tan(theta), var_type=""Float"";"
    Print #1, "assign X = sqrt(TempX**2+W**2), var_type=""Float"";"

    Print #1, "assign OriX1 = X*cos(alpha) , var_type=""Float"";"
    Print #1, "assign OriY1 = -X*sin(alpha) , var_type=""Float"";"
    Print #1, "assign OriX2 = -X*cos(alpha) , var_type=""Float"";"
    Print #1, "assign OriY2 = X*sin(alpha) , var_type=""Float"";"

    Print #1, "assign Len = OriY2*2 , var_type=""Float"";"

    Print #1, "assign p11 = OriX1+H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p12 = OriY1+H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p21 = OriX1-H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p22 = OriY1-H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p31 = OriX2-H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p32 = OriY2-H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p41 = OriX2+H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p42 = OriY2+H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p51 = OriX2 , var_type=""Float"";"
    Print #1, "assign p52 = OriY2+Len , var_type=""Float"";"

    Print #1, "assign p61 = OriX1 , var_type=""Float"";"
    Print #1, "assign p62 = OriY1-Len , var_type=""Float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p61, -p62,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "vert6 = " & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p61, -p62,"
                              Print #1, "vert3 = 0, p21, -p22,"
                              Print #1, "vert4 = 0, p31, -p32,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "vert6 = 0, p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p61, -p62,"
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "vert6 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

Public Sub VB_XBracing_YZ_Type02(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, valSP3 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Left of Top"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Point at Right of Top"";"
    Print #1, "assign end3_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign length_x1 = (end2_y - end3_y), var_type=""Float"";"
    Print #1, "assign length_z1 = (end3_z - end2_z), var_type=""Float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type=""Float"";"
' !assign alpha1 = atand(length_z1/length_x1), var_type = "float";

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !Space 3 Length"
    Print #1, "assign SP3 = " & CStr(valSP3) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !Offset"
'    Print #1, "assign OF = " & CStr(valOF) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"

    Print #1, "assign H = W+SP3+SP1+(BTB_EA*BTB_Spa)+SP2, var_type=""Float"";"

    Print #1, "assign theta = 90-2*alpha, var_type=""Float"";"

    Print #1, "assign TempX = W*tan(theta), var_type=""Float"";"
    Print #1, "assign X = sqrt(TempX**2+W**2), var_type=""Float"";"

    Print #1, "assign OriX1 = X*cos(alpha) , var_type=""Float"";"
    Print #1, "assign OriY1 = -X*sin(alpha) , var_type=""Float"";"
    Print #1, "assign OriX2 = -X*cos(alpha) , var_type=""Float"";"
    Print #1, "assign OriY2 = X*sin(alpha) , var_type=""Float"";"

    Print #1, "assign Len = OriY2*2 , var_type=""Float"";"

    Print #1, "assign p11 = OriX1+H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p12 = OriY1+H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p21 = OriX1-H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p22 = OriY1-H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p31 = OriX2-H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p32 = OriY2-H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p41 = OriX2+H*cos(alpha) , var_type=""Float"";"
    Print #1, "assign p42 = OriY2+H*sin(alpha) , var_type=""Float"";"

    Print #1, "assign p51 = OriX2 , var_type=""Float"";"
    Print #1, "assign p52 = OriY2+Len , var_type=""Float"";"

    Print #1, "assign p61 = OriX1 , var_type=""Float"";"
    Print #1, "assign p62 = OriY1-Len , var_type=""Float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0 , p11, p12,"
                              Print #1, "vert2 = 0, p61, p62,"
                              Print #1, "vert3 = 0, p21, p22,"
                              Print #1, "vert4 = 0, p31, p32,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "vert6 = 0, p41, p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p61, p62,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, p52,"
                              Print #1, "vert6 = -" & CStr(valGPT) & ", p41, p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p61, p62,"
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "vert6 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

Public Sub VB_KBracing_YZ_Type01(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Left of Bottom"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Point at Right of Bottom"";"
    Print #1, "assign end3_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign length_x1 = (end2_y - end3_y), var_type=""Float"";"
    Print #1, "assign length_z1 = (end2_z - end3_z), var_type=""Float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha1 = atand(length_z1/length_x1), var_type=""Float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type=""Float"";"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"
    Print #1, "assign BeHD = BeD/2, var_type=""Float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign TempX1 = (BeHD/sin(alpha1))+(0.015/sin(alpha1))+(W/tan(alpha1)), var_type=""Float"";"
    Print #1, "assign H1 = TempX1+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign Hsin_al1 = H1*sin(alpha1), var_type=""Float"";"
    Print #1, "assign Hcos_al1 = H1*cos(alpha1), var_type=""Float"";"

    Print #1, "assign Wsin_al1 = W*sin(alpha1), var_type=""Float"";"
    Print #1, "assign Wcos_al1 = W*cos(alpha1), var_type=""Float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p12 = BeHD, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = Hcos_al1-Wsin_al1, var_type=""Float"";"
    Print #1, "assign p42 = Hsin_al1+Wcos_al1, var_type=""Float"";"

    Print #1, "assign p51 = Hcos_al1+Wsin_al1, var_type=""Float"";"
    Print #1, "assign p52 = Hsin_al1-Wcos_al1, var_type=""Float"";"

    Print #1, "assign p61 = Hcos_al1+Wsin_al1, var_type=""Float"";"
    Print #1, "assign p62 = BeHD, var_type=""Float"";"

'    Print #1, "plc_area"
'    Print #1, "vert1 = 0, p11, -p12,"
'    Print #1, "vert2 = 0, p21, -p22,"
'    Print #1, "vert3 = 0, p31, -p32,"
'    Print #1, "vert4 = 0, -p41, -p42,"
'    Print #1, "vert5 = 0, -p51, -p52,"
'    Print #1, "vert6 = 0, -p61, -p62,"
'    Print #1, "class = " & gstr_VBClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""GP_" & CStr(valGPT) & """" & ", "
'    Print #1, "thickness = GPT;"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "vert6 = " & CStr(valGPT) & ", -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "vert6 = 0, -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "vert6 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

Public Sub VB_KBracing_YZ_Type02(valPathName As String, _
    valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
Open valPathName For Output As #1
    Print #1, "evaluate verify = ""yes"";"
    Print #1, "Default delete_log = ""yes"";"

    Print #1, "origin prompt = ""Pick Point at Left of Bottom"";"
    Print #1, "assign end1_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end1_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end1_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Center Point"";"
    Print #1, "assign end2_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end2_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end2_z=%%point_z, var_type=""Float"";"

    Print #1, "origin prompt = ""Pick Point at Right of Bottom"";"
    Print #1, "assign end3_x=%%point_x, var_type=""Float"";"
    Print #1, "assign end3_y=%%point_y, var_type=""Float"";"
    Print #1, "assign end3_z=%%point_z, var_type=""Float"";"

    Print #1, "assign length_x = (end2_y - end1_y), var_type=""Float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type=""Float"";"

    Print #1, "assign length_x1 = (end2_y - end3_y), var_type=""Float"";"
    Print #1, "assign length_z1 = (end2_z - end3_z), var_type=""Float"";"

    Print #1, "assign alpha = -1*atand(length_z/length_x), var_type=""Float"";"
    Print #1, "assign alpha1 = atand(length_z1/length_x1), var_type=""Float"";"

    Print #1, "origin local = end2_x, end2_y, end2_z;"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' !Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type=""Float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""Float"";"

    Print #1, "assign TempX = (0.015/sin(alpha))+(W/tan(alpha)), var_type=""Float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign TempX1 = (0.015/sin(alpha1))+(W/tan(alpha1)), var_type=""Float"";"
    Print #1, "assign H1 = TempX1+SP1+SP2+BTB_EA*BTB_Spa, var_type=""Float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type=""Float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type=""Float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type=""Float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type=""Float"";"

    Print #1, "assign Hsin_al1 = H1*sin(alpha1), var_type=""Float"";"
    Print #1, "assign Hcos_al1 = H1*cos(alpha1), var_type=""Float"";"

    Print #1, "assign Wsin_al1 = W*sin(alpha1), var_type=""Float"";"
    Print #1, "assign Wcos_al1 = W*cos(alpha1), var_type=""Float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p12 = 0, var_type=""Float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type=""Float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type=""Float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type=""Float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type=""Float"";"

    Print #1, "assign p41 = Hcos_al1-Wsin_al1, var_type=""Float"";"
    Print #1, "assign p42 = Hsin_al1+Wcos_al1, var_type=""Float"";"

    Print #1, "assign p51 = Hcos_al1+Wsin_al1, var_type=""Float"";"
    Print #1, "assign p52 = Hsin_al1-Wcos_al1, var_type=""Float"";"

    Print #1, "assign p61 = Hcos_al1+Wsin_al1, var_type=""Float"";"
    Print #1, "assign p62 = 0, var_type=""Float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "vert6 = " & CStr(valGPT) & ", -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "vert6 = 0, -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "vert6 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p61, -p62,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
    
Close #1
End Sub

' ----- 원본 File : GP_XZ_Column_Left_Top_offset to X(2).PML -----

Public Sub VB_Offset_RightTop_YZ_Type01(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1 * alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Depth"
'    Print #1, "assign CD = " & CStr(valCd) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
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
'    Print #1, "assign CHD = CD / 2, var_type = ""float"";"
    Print #1, "assign BeHD = BeD / 2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD / Sin(alpha)) + (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BeHD, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    ' assign Len = ((0.015/sin(alpha))+((2*W)/tan(alpha)))*cos(15), var_type = ""float"";"
    Print #1, "assign alpha2 = (90 - alpha), var_type = ""float"";"

    Print #1, "assign Len = p32-BeHd, var_type = ""float"";"
    Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"

    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = BeHD, var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1

End Sub

' ----- 원본 File : GP_XZ_Column_Left_Top_offset to X(1).PML -----

Public Sub VB_Offset_RightTop_YZ_Type02(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z / length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1 * alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Depth"
'    Print #1, "assign CD = " & CStr(valCd) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign Bed = " & CStr(valBed) & ", var_type = ""float"";"
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
'    Print #1, "assign CHD = CD / 2, var_type = ""float"";"

    Print #1, "assign TempX = (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = 0, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign P31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    ' assign Len = ((0.015/sin(alpha))+((2*W)/tan(alpha)))*cos(15), var_type = ""float"";"
    Print #1, "assign alpha2 = (90 - alpha), var_type = ""float"";"

    Print #1, "assign Len = p32, var_type = ""float"";"
    Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"

    Print #1, "assign p41 = P31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";"

    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

' ----- 원본 File : GP_XZ_Column_Right_Top_offset to X(2).PML -----

Public Sub VB_Offset_LeftTop_YZ_Type01(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Depth"
'    Print #1, "assign CD = " & CStr(valCd) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
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
'    Print #1, "assign CHD = CD/2, var_type = ""float"";"
    Print #1, "assign BeHD = BeD/2, var_type = ""float"";"

    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = BeHD, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al-Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al+Wcos_al, var_type = ""float"";"

    ' assign Len = ((0.015/sin(alpha))+((2*W)/tan(alpha)))*cos(15), var_type = ""float"";"
    Print #1, "assign alpha2 = (90-alpha), var_type = ""float"";"

    Print #1, "assign Len = p32-BeHd, var_type = ""float"";"
    Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"

    Print #1, "assign p41 = P31-TempX, var_type = ""float"";"
    Print #1, "assign p42 = BeHD, var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";"

'    Print #1, "plc_area"
'    Print #1, "vert1 = 0, -p11, -p12,"
'    Print #1, "vert2 = 0, -p21, -p22,"
'    Print #1, "vert3 = 0, -p31, -p32,"
'    Print #1, "vert4 = 0, -p41, -p42,"
'    Print #1, "class = " & gstr_VBClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""GP_" & CStr(valGPT) & """" & ", "
'    Print #1, "thickness = GPT;"
    
        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

    
Close #1
    
End Sub

' ----- 원본 File : GP_XZ_Column_Right_Top_offset to X(1).PML -----

Public Sub VB_Offset_LeftTop_YZ_Type02(valPathName As String, _
    valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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

    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"

    Print #1, "assign alpha1 = atand(length_z / length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"

    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Depth"
'    Print #1, "assign CD = " & CStr(valCd) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
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
'    Print #1, "assign CHD = CD / 2, var_type = ""float"";"

    Print #1, "assign TempX = (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX + SP1 + SP2 + BTB_EA * BTB_Spa, var_type = ""float"";"

    Print #1, "assign Hsin_al = H * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H * Cos(alpha), var_type = ""float"";"

    Print #1, "assign Wsin_al = W * Sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W * Cos(alpha), var_type = ""float"";"

    Print #1, "assign p11 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p12 = 0, var_type = ""float"";"

    Print #1, "assign p21 = Hcos_al + Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al - Wcos_al, var_type = ""float"";"

    Print #1, "assign p31 = Hcos_al - Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al + Wcos_al, var_type = ""float"";"

    ' assign Len = ((0.015/sin(alpha))+((2*W)/tan(alpha)))*cos(15), var_type = ""float"";"
    Print #1, "assign alpha2 = (90 - alpha), var_type = ""float"";"

    Print #1, "assign Len = p32, var_type = ""float"";"
    Print #1, "assign TempX = Len*tan(alpha2), var_type = ""float"";"

    Print #1, "assign p41 = p31 - TempX, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    ' assign p42 = P32-TempZZ, var_type = ""float"";"

    ' assign p51 = P31-TempXX, var_type = ""float"";"
    ' assign p52 = BeD, var_type = ""float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
    
End Sub

Public Sub VB_WEAK_RightBottJoint_YZ(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(30), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = 0, var_type = ""float"";"
                   
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select


Close #1
End Sub

Public Sub VB_WEAK_RightBott_YZ(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
    
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
    
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
    
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"

    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(15), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = -1*Len*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = -1*Len*cos(alpha), var_type = ""float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub
Public Sub VB_WEAK_RightBott_YZ_Type03(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
    
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
    
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
    
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"

    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(15), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = -end1_z, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = -end1_z, var_type = ""float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", p21, p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, p12,"
                              Print #1, "vert2 = 0, p21, p22,"
                              Print #1, "vert3 = 0, p31, p32,"
                              Print #1, "vert4 = 0, p41, p42,"
                              Print #1, "vert5 = 0, p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", p11, p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", p21, p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", p31, p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", p41, p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_WEAK_LeftBottJoint_YZ(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(30), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    Print #1, ""
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    Print #1, ""
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = 0, var_type = ""float"";"

    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_WEAK_LeftBott_YZ(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(15), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = -1*Len*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = -1*Len*cos(alpha), var_type = ""float"";"
               
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub
Public Sub VB_WEAK_LeftBott_YZ_Type03(valPathName As String, _
    valWD As Single, valWt As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Width"
    Print #1, "assign WD = " & CStr(valWD) & ", var_type = ""float"";"
    ' Column Flange"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign WHD = WD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (WHD/cos(alpha))+(0.015/cos(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    Print #1, "assign Len = ((W*tan(alpha))+H-(WHD/cos(alpha)))*sin(15), var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p41 = (Len*sin(alpha))+WHD, var_type = ""float"";"
    Print #1, "assign p42 = -end1_z, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = -end1_z, var_type = ""float"";"
               
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, p12,"
                              Print #1, "vert2 = 0, -p21, p22,"
                              Print #1, "vert3 = 0, -p31, p32,"
                              Print #1, "vert4 = 0, -p41, p42,"
                              Print #1, "vert5 = 0, -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p11, p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p21, p22,"
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p31, p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p41, p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", -p51, p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_WEAK_RightTop_YZ_Type01(valPathName As String, _
    valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
    
    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = BeD, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = BeD, var_type = ""float"";"
    
    
        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_WEAK_RightTop_YZ_Type02(valPathName As String, _
    valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BeHD = BeD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = BeHD, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = BeHD, var_type = ""float"";"
        
        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_WEAK_RightTop_YZ_Type03(valPathName As String, _
    valCD As Single, valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = -1*alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange width"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type = ""float"";"
    ' Column web THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
    
    Print #1, "assign TempX = (CD/(sin(alpha)*2))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = 0, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    
        Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, p11, -p12,"
                              Print #1, "vert2 = 0, p21, -p22,"
                              Print #1, "vert3 = 0, p31, -p32,"
                              Print #1, "vert4 = 0, p41, -p42,"
                              Print #1, "vert5 = 0, p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(valGPT) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(valGPT) & ", p21, -p22,"
                              Print #1, "vert3 = -" & CStr(valGPT) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(valGPT) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(valGPT) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p11, -p12,"
                              Print #1, "vert2 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p21, -p22, "
                              Print #1, "vert3 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p31, -p32,"
                              Print #1, "vert4 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p41, -p42,"
                              Print #1, "vert5 = -" & CStr(Format(valGPT / 2, "0.000")) & ", p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select

Close #1
End Sub

Public Sub VB_WEAK_LeftTop_YZ_Type01(valPathName As String, _
    valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
    
    Print #1, "assign TempX = (BeD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = BeD, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = BeD, var_type = ""float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_WEAK_LeftTop_YZ_Type02(valPathName As String, _
    valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign BeHD = BeD/2, var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
               
    Print #1, "assign TempX = (BeHD/sin(alpha))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = BeHD, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = BeHD, var_type = ""float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub

Public Sub VB_WEAK_LeftTop_YZ_Type03(valPathName As String, _
    valCD As Single, valWt As Single, valBed As Single, valBd As Single, valSP1 As Single, valSP2 As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single, valGPT As Single, valBr_Flag As String)
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
               
    Print #1, "assign length_x = (end2_y - end1_y), var_type = ""float"";"
    Print #1, "assign length_z = (end2_z - end1_z), var_type = ""float"";"
               
    Print #1, "assign alpha1 = atand(length_z/length_x), var_type = ""float"";"
    Print #1, "assign alpha = alpha1, var_type = ""float"";"
               
    Print #1, "origin local = end1_x, end1_y, end1_z;"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    ' Column Flange widht"
    Print #1, "assign CD = " & CStr(valCD) & ", var_type = ""float"";"
    ' Column web THK"
    Print #1, "assign Ft = " & CStr(valWt) & ", var_type = ""float"";"
    ' Beam Depth"
    Print #1, "assign BeD = " & CStr(valBed) & ", var_type = ""float"";"
    ' Bracing Depth"
    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    ' !Space 1 Length"
    Print #1, "assign SP1 = " & CStr(valSP1) & ", var_type=""Float"";"
    ' !Space 2 Length"
    Print #1, "assign SP2 = " & CStr(valSP2) & ", var_type=""Float"";"
    ' !BTB Space EA"
    Print #1, "assign BTB_EA = " & CStr(valBTB_EA) & ", var_type=""Float"";"
    ' !BTB Space"
    Print #1, "assign BTB_Spa = " & CStr(valBTB_Spa) & ", var_type=""Float"";"
    ' !Gusset Plate Thickness"
    Print #1, "assign GPT = " & CStr(valGPT) & ", var_type=""Float"";"
    ' !!!!!!!!!!!!!!Data input!!!!!!!!!!!!!!!!!!"
    Print #1, "assign W = BD/2 ,var_type = ""float"";"
    Print #1, "assign FHt = Ft/2, var_type = ""float"";"
    
    Print #1, "assign TempX = (CD/(sin(alpha)*2))+(0.015/sin(alpha))+(W/tan(alpha)), var_type = ""float"";"
    Print #1, "assign H = TempX+SP1+SP2+BTB_EA*BTB_Spa, var_type = ""float"";"
    
    Print #1, "assign Hsin_al = H*sin(alpha), var_type = ""float"";"
    Print #1, "assign Hcos_al = H*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign Wsin_al = W*sin(alpha), var_type = ""float"";"
    Print #1, "assign Wcos_al = W*cos(alpha), var_type = ""float"";"
    
    Print #1, "assign p11 = FHt, var_type = ""float"";"
    Print #1, "assign p12 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p21 = Hcos_al-Wsin_al, var_type = ""float"";"
    Print #1, "assign p22 = Hsin_al+Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p31 = Hcos_al+Wsin_al, var_type = ""float"";"
    Print #1, "assign p32 = Hsin_al-Wcos_al, var_type = ""float"";"
    
    Print #1, "assign p51 = FHt, var_type = ""float"";"
    Print #1, "assign p52 = 0, var_type = ""float"";"
    
    Print #1, "assign Len = P32-P52, var_type = ""float"";"
    Print #1, "assign TempXX = Len*tan(90-alpha), var_type = ""float"";"
    
    Print #1, "assign p41 = P31-TempXX, var_type = ""float"";"
    Print #1, "assign p42 = 0, var_type = ""float"";"
    
    Select Case valBr_Flag
               Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(valGPT) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(valGPT) & ", -p21, -p22,"
                              Print #1, "vert3 = " & CStr(valGPT) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(valGPT) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(valGPT) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                              Print #1, "plc_area"
                              Print #1, "vert1 = 0, -p11, -p12,"
                              Print #1, "vert2 = 0, -p21, -p22,"
                              Print #1, "vert3 = 0, -p31, -p32,"
                              Print #1, "vert4 = 0, -p41, -p42,"
                              Print #1, "vert5 = 0, -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
               Case "DAopt01", "DCopt03"
                              Print #1, "plc_area"
                              Print #1, "vert1 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p11, -p12,"
                              Print #1, "vert2 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p21, -p22, "
                              Print #1, "vert3 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p31, -p32,"
                              Print #1, "vert4 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p41, -p42,"
                              Print #1, "vert5 = " & CStr(Format(valGPT / 2, "0.000")) & ", -p51, -p52,"
                              Print #1, "class = " & gstr_VBClass & ", " & _
                                        "grade = """ & gstr_Grade & """, " & _
                                        "material = """ & gstr_Material & """, " & _
                                        "name = ""GP_" & CStr(valGPT) & """" & ", "
                              Print #1, "thickness = GPT;"
    End Select
Close #1
End Sub
Public Sub Ver_HTB_Nut_M01_YZ(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
                              
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M02_YZ(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 =  " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M03_YZ(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X2,  -HTB_Y1,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X3,  -HTB_Y1,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X4,  -HTB_Y1,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X5,  -HTB_Y1,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X6,  -HTB_Y1,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M04_YZ(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M05_YZ_2(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X2,  -HTB_Y1,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X3,  -HTB_Y1,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X4,  -HTB_Y1,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X5,  -HTB_Y1,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X6,  -HTB_Y1,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub

Public Sub Ver_HTB_Nut_M06_YZ_2(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1
    Print #1, "assign TempX = (0.015 / Sin(alpha)) + (W / Tan(alpha)), var_type = ""float"";"
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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub

Public Sub Ver_HTB_Nut_M07_YZ_1(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                           valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                                           valHTBSpace As Single, valGPThk As Single, valGage As Single, _
                                                                           valBwt As Single, valBr_Flag As String)

Dim i As Integer

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
                              
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, - HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, - HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, - HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M07_YZ_2(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                           valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                                           valHTBSpace As Single, valGPThk As Single, valGage As Single, _
                                                                           valBwt As Single, valBr_Flag As String)

Dim i As Integer

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
                              
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, - HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             
                                             If valBoltType = "II" Then
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, - HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, - HTB_Y1,"
'                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  -HTB_Y11,"
'                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1,  -HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             If valBoltType = "II" Then
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11,  -HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2 + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M07_YZ_3(valPathName As String, valBoltType As String, valBoltName As String, _
                                                                           valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                                           valHTBSpace As Single, valGPThk As Single, valGage As Single, _
                                                                           valBwt As Single, valBr_Flag As String)

Dim i As Integer

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
                              
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, - HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, - HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, - HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
'                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
'                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X1,  HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X2,  HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X3,  HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X4,  HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X5,  HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X6,  HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1,  HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2,  HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3,  HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4,  HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5,  HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6,  HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                             Print #1, "plc_area"
'                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X1,  -HTB_Y1,"
'                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X2,  -HTB_Y2,"
'                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X3,  -HTB_Y3,"
'                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X4,  -HTB_Y4,"
'                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X5,  -HTB_Y5,"
'                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X6,  -HTB_Y6,"
'                                             Print #1, "class = " & gstr_VBClass & ", " & _
'                                                       "grade = """ & gstr_Grade & """, " & _
'                                                       "material = """ & gstr_Material & """, " & _
'                                                       "name = ""HTB_" & valBoltName & """, "
'                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X11,  HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X21,  HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X31,  HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X41,  HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X51,  HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", -HTB_X61,  HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                                            
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11,  HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21,  HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31,  HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41,  HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51,  HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61,  HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
'
'                                                            Print #1, "plc_area"
'                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X11,  -HTB_Y11,"
'                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X21,  -HTB_Y21,"
'                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X31,  -HTB_Y31,"
'                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X41,  -HTB_Y41,"
'                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X51,  -HTB_Y51,"
'                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", -HTB_X61,  -HTB_Y61,"
'                                                            Print #1, "class = " & gstr_VBClass & ", " & _
'                                                                      "grade = """ & gstr_Grade & """, " & _
'                                                                      "material = """ & gstr_Material & """, " & _
'                                                                      "name = ""HTB_" & valBoltName & """, "
'                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub
Public Sub Ver_HTB_Nut_M08_YZ_1(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, ByVal valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk + valNutHei, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X2,  -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X3,  -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X4,  -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X5,  -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X6,  -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X1,  -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X2,  -HTB_Y1,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X3,  -HTB_Y1,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X4,  -HTB_Y1,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X5,  -HTB_Y1,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X6,  -HTB_Y1,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X11,  -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X21,  -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X31,  -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X41,  -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X51,  -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", - HTB_X61,  -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub

Public Sub Ver_HTB_Nut_M08_YZ_2(valPathName As String, valBoltType As String, valBoltName As String, valNutDia As Single, valNutHei As Single, valHTBNum As Integer, _
                                                  valHTBSpace As Single, valGPThk As Single, valGage As Single, valBwt As Single, valBr_Flag As String)

Dim i As Integer

'If valBoltType = "I" Then valBoltName = valBoltName & "-1"
'If valBoltType = "II" Then valBoltName = valBoltName & "-2"

Open valPathName For Append As #1

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
                                             Print #1, "assign ceta1 = alpha1+ atand(" & CStr(valGage) & "/HH), var_type = ""float"";"
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
               
               Select Case valBr_Flag
                              Case "Aopt01", "Copt03", "DAopt06", "Topt01", "Topt06"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                                             
                              Case "Aopt02", "Copt04", "DAopt05", "Topt05"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = -" & CStr(Format(valGPThk + valNutHei, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
                              Case "DAopt01", "DCopt03"
                                             Print #1, "plc_area"
                                             Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X1, -HTB_Y1,"
                                             Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X2, -HTB_Y2,"
                                             Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X3, -HTB_Y3,"
                                             Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X4, -HTB_Y4,"
                                             Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X5, -HTB_Y5,"
                                             Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X6, -HTB_Y6,"
                                             Print #1, "class = " & gstr_VBClass & ", " & _
                                                       "grade = """ & gstr_Grade & """, " & _
                                                       "material = """ & gstr_Material & """, " & _
                                                       "name = ""HTB_" & valBoltName & """, "
                                             Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             
                                             If valBoltType = "II" Then
                                                            Print #1, "plc_area"
                                                            Print #1, "vert1 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X11, -HTB_Y11,"
                                                            Print #1, "vert2 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X21, -HTB_Y21,"
                                                            Print #1, "vert3 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X31, -HTB_Y31,"
                                                            Print #1, "vert4 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X41, -HTB_Y41,"
                                                            Print #1, "vert5 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X51, -HTB_Y51,"
                                                            Print #1, "vert6 = " & CStr(Format(valGPThk / 2, "0.000")) & ", HTB_X61, -HTB_Y61,"
                                                            Print #1, "class = " & gstr_VBClass & ", " & _
                                                                      "grade = """ & gstr_Grade & """, " & _
                                                                      "material = """ & gstr_Material & """, " & _
                                                                      "name = ""HTB_" & valBoltName & """, "
                                                            Print #1, "thickness = " & CStr(valNutHei) & ";"
                                             End If
               End Select
    Next i
    
Close #1
End Sub


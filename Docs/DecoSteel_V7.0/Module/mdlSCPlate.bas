Attribute VB_Name = "mdlSCPlate"
Public Sub SC_PML(valPath As String, ByVal valJobName As String, ByVal valCode As String, _
                                                ByVal valFormCode As String, valDir As String, valSCP_Flag As String, _
                                                valType As String, valPlateType As String, valUnit As String, _
                                                valColumn As String, valBeam As String, valSubBeam As String)
Dim TempX1 As Single, TempX2 As Single, TempY1 As Single, TempY2 As Single, TempThk As Single, TempStiffThk As Single
Dim TempBoltDia As Single, TempBoltEA As Integer, TempNutDia As Single, TempNutHei As Single, _
    TempName As String, TempGap As Single, TempHTB_Space As Single
Dim TempA As Single, TempB As Single, TempC As Single, TempD As Single, TempE As Single
Dim TempCdepth As Single, TempCwidth As Single, TempCft As Single, TempCwt As Single
Dim TempBedepth As Single, TempBewidth As Single, TempBewt As Single, TempBeft As Single
Dim TempSubBedepth As Single, TempSubBewidth As Single, TempSubBewt As Single, TempSubBeft As Single
    
    Call SCData_Call(valJobName, valCode, valFormCode, valColumn, valBeam, valSubBeam, valDir, valSCP_Flag)
    
    TempThk = convert(gsin_Pthk, gstr_Unit, valUnit)
    TempStiffThk = convert(gsin_StiffThk, gstr_Unit, valUnit)
    TempHTB_Space = convert(gsin_HTB_Space, gstr_Unit, valUnit)
    
    TempCdepth = convert(gsin_Cdepth, "mm", valUnit)
    TempCwidth = convert(gsin_Cwidth, "mm", valUnit)
    TempCft = convert(gsin_Cft, "mm", valUnit)
    TempCwt = convert(gsin_Cwt, "mm", valUnit)
    TempBedepth = convert(gsin_Bedepth, "mm", valUnit)
    TempBewidth = convert(gsin_Bewidth, "mm", valUnit)
    TempBewt = convert(gsin_Bewt, "mm", valUnit)
    TempBeft = convert(gsin_Beft, "mm", valUnit)
    TempSubBedepth = convert(gsin_SubBedepth, "mm", valUnit)
    TempSubBewidth = convert(gsin_SubBewidth, "mm", valUnit)
    TempSubBewt = convert(gsin_SubBewt, "mm", valUnit)
    TempSubBeft = convert(gsin_SubBeft, "mm", valUnit)
    
    TempGap = convert(gsin_Gap, gstr_Unit, valUnit)
    TempA = convert(gsin_A, gstr_Unit, valUnit)
    TempB = convert(gsin_B, gstr_Unit, valUnit)
    TempC = convert(gsin_C, gstr_Unit, valUnit)
    TempD = convert(gsin_D, gstr_Unit, valUnit)
    TempE = convert(gsin_E, gstr_Unit, valUnit)
    TempBoltDia = convert(gsin_BoltDia, gstr_NutUnit, valUnit)
    TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
    TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)
    
    Select Case valDir
        Case "X"
            Select Case valSCP_Flag
                Case "Module01"
                    Select Case valType
                        Case "Type01"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module01_XZ_Type01_A(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module01_XZ_Type01_B(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                        Case "Type02"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module01_XZ_Type02_A(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module01_XZ_Type02_B(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                    End Select
                Case "Module02"
                    Select Case valType
                        Case "Type01"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module02_XZ_Type01_A(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module02_XZ_Type01_B(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type C"
                                    Call SC_Module02_XZ_Type01_C(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type D"
                                    Call SC_Module02_XZ_Type01_D(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                        Case "Type02"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module02_XZ_Type02_A(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module02_XZ_Type02_B(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type C"
                                    Call SC_Module02_XZ_Type02_C(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type D"
                                    Call SC_Module02_XZ_Type02_D(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                    End Select
            End Select
        Case "Y"
            Select Case valSCP_Flag
                Case "Module01"
                    Select Case valType
                        Case "Type01"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module01_YZ_Type01_A(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module01_YZ_Type01_B(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                        Case "Type02"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module01_YZ_Type02_A(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module01_YZ_Type02_B(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                    End Select
                Case "Module02"
                    Select Case valType
                        Case "Type01"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module02_YZ_Type01_A(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module02_YZ_Type01_B(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type C"
                                    Call SC_Module02_YZ_Type01_C(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type D"
                                    Call SC_Module02_YZ_Type01_D(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                        Case "Type02"
                            Select Case valPlateType
                                Case "Type A"
                                    Call SC_Module02_YZ_Type02_A(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type B"
                                    Call SC_Module02_YZ_Type02_B(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempStiffThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type C"
                                    Call SC_Module02_YZ_Type02_C(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                                Case "Type D"
                                    Call SC_Module02_YZ_Type02_D(valPath, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempSubBedepth, TempSubBewidth, TempSubBewt, TempSubBeft, _
                                        TempThk, TempGap, _
                                        TempA, TempB, TempC, TempD, TempE, gsin_HTB_SNum, TempHTB_Space)
                            End Select
                    End Select
            End Select
    End Select

    Select Case valSCP_Flag
        Case "Module01"
            Call PrintNut_Shear(valPath, TempThk, TempCwidth, TempCwt, TempBedepth, TempBewt, TempGap, _
                                                   TempA, TempB, TempC, TempD, TempE, TempBoltDia, TempNutDia, _
                                                   TempNutHei, gstr_BoltName, TempHTB_Space, gin_BoltEA, valPlateType, valType, gstr_Type, valDir)
        Case "Module02"
            Call PrintNut_Shear(valPath, TempThk, TempBewidth, TempBewt, TempSubBedepth, TempSubBewt, TempGap, _
                                                   TempA, TempB, TempC, TempD, TempE, TempBoltDia, TempNutDia, TempNutHei, _
                                                   gstr_BoltName, TempHTB_Space, gin_BoltEA, valPlateType, valType, gstr_Type, valDir)
    End Select
End Sub

Public Sub SC_Module01_XZ_Type01_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valSTt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, 0,"
    Print #1, "vert2 = Px2, Py, 0,"
    Print #1, "vert3 = Px2, Py, -Pz1,"
    Print #1, "vert4 = Px3, Py, -Pz1,"
    Print #1, "vert5 = Px3, Py, -Pz2,"
    Print #1, "vert6 = Px2, Py, -Pz2,"
    Print #1, "vert7 = Px2, Py, -Pz,"
    Print #1, "vert8 = Px1, Py, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spx2, spy, -SPz,"
    Print #1, "vert2 = spx1, spy, -SPz,"
    Print #1, "vert3 = spx1, -spy, -SPz,"
    Print #1, "vert4 = spx2, -spy, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module01_XZ_Type01_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valSTt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt =  " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, 0,"
    Print #1, "vert2 = Px2, Py, 0,"
    Print #1, "vert3 = Px2, Py, -Pz1,"
    Print #1, "vert4 = Px3, Py, -Pz1,"
    Print #1, "vert5 = Px3, Py, -Pz2,"
    Print #1, "vert6 = Px2, Py, -Pz2,"
    Print #1, "vert7 = Px2, Py, -Pz,"
    Print #1, "vert8 = Px1, Py, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spx2, spy, -SPz,"
    Print #1, "vert2 = spx1, spy, -SPz,"
    Print #1, "vert3 = spx1, -spy, -SPz,"
    Print #1, "vert4 = spx2, -spy, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module01_XZ_Type02_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valSTt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, 0,"
    Print #1, "vert2 = -Px2, Py, 0,"
    Print #1, "vert3 = -Px2, Py, -Pz1,"
    Print #1, "vert4 = -Px3, Py, -Pz1,"
    Print #1, "vert5 = -Px3, Py, -Pz2,"
    Print #1, "vert6 = -Px2, Py, -Pz2,"
    Print #1, "vert7 = -Px2, Py, -Pz,"
    Print #1, "vert8 = -Px1, Py, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -spx1, spy, -SPz,"
    Print #1, "vert2 = -spx2, spy, -SPz,"
    Print #1, "vert3 = -spx2, -spy, -SPz,"
    Print #1, "vert4 = -spx1, -spy, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module01_XZ_Type02_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valSTt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, 0,"
    Print #1, "vert2 = -Px2, Py, 0,"
    Print #1, "vert3 = -Px2, Py, -Pz1,"
    Print #1, "vert4 = -Px3, Py, -Pz1,"
    Print #1, "vert5 = -Px3, Py, -Pz2,"
    Print #1, "vert6 = -Px2, Py, -Pz2,"
    Print #1, "vert7 = -Px2, Py, -Pz,"
    Print #1, "vert8 = -Px1, Py, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -spx1, spy, -SPz,"
    Print #1, "vert2 = -spx2, spy, -SPz,"
    Print #1, "vert3 = -spx2, -spy, -SPz,"
    Print #1, "vert4 = -spx1, -spy, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type01_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, -Pz1,"
    Print #1, "vert2 = Px2, Py, -Pz1,"
    Print #1, "vert3 = Px2, Py, -Pz2,"
    Print #1, "vert4 = Px1, Py, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type01_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = (cw/2 - cwt/2)+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, -Pz1,"
    Print #1, "vert2 = Px2, Py, -Pz1,"
    Print #1, "vert3 = Px2, Py, -Pz2,"
    Print #1, "vert4 = Px1, Py, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type01_C(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, -Pz1,"
    Print #1, "vert2 = Px2, Py, -Pz1,"
    Print #1, "vert3 = Px2, Py, -Pz2,"
    Print #1, "vert4 = Px3, Py, -Pz3,"
    Print #1, "vert4 = Px1, Py, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"
Close #1
End Sub

Public Sub SC_Module02_XZ_Type01_D(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py, -Pz1,"
    Print #1, "vert2 = Px2, Py, -Pz1,"
    Print #1, "vert3 = Px2, Py, -Pz2,"
    Print #1, "vert4 = Px3, Py, -Pz3,"
    Print #1, "vert4 = Px1, Py, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type02_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"


    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"


    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, -Pz1,"
    Print #1, "vert2 = -Px2, Py, -Pz1,"
    Print #1, "vert3 = -Px2, Py, -Pz2,"
    Print #1, "vert4 = -Px1, Py, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type02_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
     Print #1, "Default delete_log = ""yes"";"
   'ONLY RIGHT BEAM EXISTING
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"


    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"


    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, -Pz1,"
    Print #1, "vert2 = -Px2, Py, -Pz1,"
    Print #1, "vert3 = -Px2, Py, -Pz2,"
    Print #1, "vert4 = -Px1, Py, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_XZ_Type02_C(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, -Pz1,"
    Print #1, "vert2 = -Px2, Py, -Pz1,"
    Print #1, "vert3 = -Px2, Py, -Pz2,"
    Print #1, "vert4 = -Px3, Py, -Pz3,"
    Print #1, "vert4 = -Px1, Py, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"
Close #1
End Sub

Public Sub SC_Module02_XZ_Type02_D(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py, -Pz1,"
    Print #1, "vert2 = -Px2, Py, -Pz1,"
    Print #1, "vert3 = -Px2, Py, -Pz2,"
    Print #1, "vert4 = -Px3, Py, -Pz3,"
    Print #1, "vert4 = -Px1, Py, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


Close #1
End Sub

Public Sub SC_Module01_YZ_Type01_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd-STt, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, 0,"
    Print #1, "vert2 = Py, -Px2, 0,"
    Print #1, "vert3 = Py, -Px2, -Pz1,"
    Print #1, "vert4 = Py, -Px3, -Pz1,"
    Print #1, "vert5 = Py, -Px3, -Pz2,"
    Print #1, "vert6 = Py, -Px2, -Pz2,"
    Print #1, "vert7 = Py, -Px2, -Pz,"
    Print #1, "vert8 = Py, -Px1, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spy, -spx1, -SPz,"
    Print #1, "vert2 = spy, -spx2, -SPz,"
    Print #1, "vert3 = -spy, -spx2, -SPz,"
    Print #1, "vert4 = -spy, -spx1, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module01_YZ_Type01_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd-STt, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, 0,"
    Print #1, "vert2 = Py, -Px2, 0,"
    Print #1, "vert3 = Py, -Px2, -Pz1,"
    Print #1, "vert4 = Py, -Px3, -Pz1,"
    Print #1, "vert5 = Py, -Px3, -Pz2,"
    Print #1, "vert6 = Py, -Px2, -Pz2,"
    Print #1, "vert7 = Py, -Px2, -Pz,"
    Print #1, "vert8 = Py, -Px1, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spy, -spx1, -SPz,"
    Print #1, "vert2 = spy, -spx2, -SPz,"
    Print #1, "vert3 = -spy, -spx2, -SPz,"
    Print #1, "vert4 = -spy, -spx1, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module01_YZ_Type02_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2+PT, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd-STt, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, 0,"
    Print #1, "vert2 = Py, Px2, 0,"
    Print #1, "vert3 = Py, Px2, -Pz1,"
    Print #1, "vert4 = Py, Px3, -Pz1,"
    Print #1, "vert5 = Py, Px3, -Pz2,"
    Print #1, "vert6 = Py, Px2, -Pz2,"
    Print #1, "vert7 = Py, Px2, -Pz,"
    Print #1, "vert8 = Py, Px1, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spy, spx2, -SPz,"
    Print #1, "vert2 = spy, spx1, -SPz,"
    Print #1, "vert3 = -spy, spx1, -SPz,"
    Print #1, "vert4 = -spy, spx2, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"


Close #1
End Sub

Public Sub SC_Module01_YZ_Type02_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cw/2, var_type = ""float"";"
    Print #1, "assign Py = Bwt/2+PT, var_type = ""float"";"
    Print #1, "assign Pz = Bd, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px3 = Px2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPX2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPY = SPl/2, var_type = ""float"";"
    Print #1, "assign SPz = Bd-STt, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, 0,"
    Print #1, "vert2 = Py, Px2, 0,"
    Print #1, "vert3 = Py, Px2, -Pz1,"
    Print #1, "vert4 = Py, Px3, -Pz1,"
    Print #1, "vert5 = Py, Px3, -Pz2,"
    Print #1, "vert6 = Py, Px2, -Pz2,"
    Print #1, "vert7 = Py, Px2, -Pz,"
    Print #1, "vert8 = Py, Px1, -Pz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


    'STIFFNER PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = spy, spx2, -SPz,"
    Print #1, "vert2 = spy, spx1, -SPz,"
    Print #1, "vert3 = -spy, spx1, -SPz,"
    Print #1, "vert4 = -spy, spx2, -SPz,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type01_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, -Pz1,"
    Print #1, "vert2 = Py, -Px2, -Pz1,"
    Print #1, "vert3 = Py, -Px2, -Pz2,"
    Print #1, "vert4 = Py, -Px1, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


Close #1
End Sub

Public Sub SC_Module02_YZ_Type01_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, -Pz1,"
    Print #1, "vert2 = Py, -Px2, -Pz1,"
    Print #1, "vert3 = Py, -Px2, -Pz2,"
    Print #1, "vert4 = Py, -Px1, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type01_C(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, -Pz1,"
    Print #1, "vert2 = Py, -Px2, -Pz1,"
    Print #1, "vert3 = Py, -Px2, -Pz2,"
    Print #1, "vert4 = Py, -Px3, -Pz3,"
    Print #1, "vert4 = Py, -Px1, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"


Close #1
End Sub

Public Sub SC_Module02_YZ_Type01_D(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px1, -Pz1,"
    Print #1, "vert2 = Py, -Px2, -Pz1,"
    Print #1, "vert3 = Py, -Px2, -Pz2,"
    Print #1, "vert4 = Py, -Px3, -Pz3,"
    Print #1, "vert4 = Py, -Px1, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type02_A(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+PT, var_type = ""float"";"


    Print #1, "assign TempPz = Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, -Pz1,"
    Print #1, "vert2 = Py, Px2, -Pz1,"
    Print #1, "vert7 = Py, Px2, -Pz2,"
    Print #1, "vert8 = Py, Px1, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type02_B(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, valSTt As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+PT, var_type = ""float"";"


    Print #1, "assign TempPz = Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign TempPz1 = (Bd-(Bft*2)-TempPz)/2, var_type = ""float"";"

    Print #1, "assign Px2 = cw/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Pz1 = Bft+TempPz1, var_type = ""float"";"
    Print #1, "assign Pz2 = Pz1+TempPz, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, -Pz1,"
    Print #1, "vert2 = Py, Px2, -Pz1,"
    Print #1, "vert7 = Py, Px2, -Pz2,"
    Print #1, "vert8 = Py, Px1, -Pz2,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type02_C(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Aval+Bval+Eval, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, -Pz1,"
    Print #1, "vert2 = Py, Px2, -Pz1,"
    Print #1, "vert3 = Py, Px2, -Pz2,"
    Print #1, "vert4 = Py, Px3, -Pz3,"
    Print #1, "vert4 = Py, Px1, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"

Close #1
End Sub

Public Sub SC_Module02_YZ_Type02_D(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valPt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBTB_EA As Integer, valBTB_Spa As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXISTING"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    Print #1, "assign Bd = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign Bw = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign Pt = " & CStr(valPt) & ", var_type = ""float"";"

    Print #1, "assign Gap = " & CStr(valGap) & ", var_type = ""float"";"
    Print #1, "assign Aval = " & CStr(valA) & ", var_type = ""float"";"
    Print #1, "assign Bval = " & CStr(valB) & ", var_type = ""float"";"
    Print #1, "assign Cval = " & CStr(valC) & ", var_type = ""float"";"
    Print #1, "assign Dval = " & CStr(valD) & ", var_type = ""float"";"
    Print #1, "assign Eval = " & CStr(valE) & ", var_type = ""float"";"
    Print #1, "assign BtoB_Spa = " & CStr(valBTB_Spa) & ", var_type = ""float"";"
    Print #1, "assign Bolt_Spa_EA = " & CStr(valBTB_EA) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px1 = cwt/2, var_type = ""float"";"
    Print #1, "assign Px2 = cwt/2+Gap+Cval+Dval+BtoB_Spa*Bolt_Spa_EA, var_type = ""float"";"
    Print #1, "assign Px3 = cw/2, var_type = ""float"";"

    Print #1, "assign Py = Bwt/2+Pt, var_type = ""float"";"

    Print #1, "assign Pz1 = Bft, var_type = ""float"";"
    Print #1, "assign Pz2 = Bd, var_type = ""float"";"
    Print #1, "assign Pz3 = Cd-Cft, var_type = ""float"";"

    'PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px1, -Pz1,"
    Print #1, "vert2 = Py, Px2, -Pz1,"
    Print #1, "vert3 = Py, Px2, -Pz2,"
    Print #1, "vert4 = Py, Px3, -Pz3,"
    Print #1, "vert4 = Py, Px1, -Pz3,"
    Print #1, "class = " & gstr_SCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""BP_" & CStr(valPt) & """" & ", "
    Print #1, "thickness = PT;"
Close #1

End Sub

Public Sub PrintNut_Shear(valPathName As String, _
    valPt As Single, valCWidth As Single, valCwt As Single, valBdepth As Single, valBwt As Single, valGap As Single, _
    valA As Single, valB As Single, valC As Single, valD As Single, valE As Single, _
    valBoltDia As Single, valNutDia As Single, valNutHei As Single, valBoltName As String, _
    valBtoBSpace As Single, valBoltEA As Integer, valType As String, valRL_Flag As String, valBoltType_Flag As String, _
    valDir_Flag As String)


Dim Seta(1 To 6) As Single
Dim Nx() As Single, Ny() As Single
Dim BX() As Single, By() As Single
Dim i As Integer, j As Integer, k As Integer
Dim Plate_Depth As Single, Plate_Space As Single
Dim RL_Flag  As String
Dim BSpace_EA As Integer

For i = 1 To 6
    If i = 1 Then
        Seta(1) = 60
    Else
        Seta(i) = 60 + Seta(i - 1)
    End If
Next i
If valBoltType_Flag = "II" Then
               BSpace_EA = valBoltEA / 2 - 1
Else
               BSpace_EA = valBoltEA - 1
End If
Select Case valType
    Case "Type A"
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        Plate_Depth = valC + valD + valBtoBSpace * BSpace_EA
        Plate_Space = (valBdepth - Plate_Depth) / 2
        
        If valBoltType_Flag = "I" Then
               For i = 1 To valBoltEA
                      BX(i) = valCWidth / 2 + valGap + valA
                       If i = 1 Then
                                     By(1) = -Plate_Space - valC
                      Else
                                     By(i) = By(1) - valBtoBSpace * (i - 1)
                      End If
               Next i
        Else
               k = 2
              For i = 1 To valBoltEA / 2
                      BX(i) = valCWidth / 2 + valGap + valA
                      If i = 1 Then
                                     By(1) = -Plate_Space - valC
                      Else
                                     By(i) = By(1) - valBtoBSpace * (k - 1)
                                     k = k + 1
                      End If
               Next i
               k = 2
               For i = valBoltEA / 2 + 1 To valBoltEA
                      BX(i) = valCWidth / 2 + valGap + valA + valE
                      If i = valBoltEA / 2 + 1 Then
                                     By(i) = -Plate_Space - valC
                      Else
                                     By(i) = By(valBoltEA / 2 + 1) - valBtoBSpace * (k - 1)
                                     k = k + 1
                      End If
               Next i
        End If
    Case "Type B"
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
         If valBoltType_Flag = "I" Then
               Plate_Depth = valA + valB
               Plate_Space = (valBdepth - Plate_Depth) / 2
               For i = 1 To valBoltEA
                              By(i) = -Plate_Space - valA
                              If i = 1 Then
                                             BX(1) = valCWidth / 2 + valGap + valC
                              Else
                                             BX(i) = BX(1) + valBtoBSpace * (i - 1)
                              End If
               Next i
         Else
               Plate_Depth = valA + valB + valE
               Plate_Space = (valBdepth - Plate_Depth) / 2
               k = 2
               For i = 1 To valBoltEA / 2
                              By(i) = -Plate_Space - valA
                              If i = 1 Then
                                             BX(1) = valCWidth / 2 + valGap + valC
                              Else
                                             BX(i) = BX(1) + valBtoBSpace * (k - 1)
                                             k = k + 1
                              End If
               Next i
               k = 2
               For i = valBoltEA / 2 + 1 To valBoltEA
                              By(i) = -Plate_Space - valA - valE
                              If i = valBoltEA / 2 + 1 Then
                                             BX(i) = valCWidth / 2 + valGap + valC
                              Else
                                             BX(i) = BX(valBoltEA / 2 + 1) + valBtoBSpace * (k - 1)
                                             k = k + 1
                              End If
               Next i
         End If
    Case "Type C"
        
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        Plate_Depth = valC + valD + valBtoBSpace * BSpace_EA
        Plate_Space = (valBdepth - Plate_Depth) / 2
        If valBoltType_Flag = "I" Then
               For i = 1 To valBoltEA
                            BX(i) = valGap + valA
                            If i = 1 Then
                                             By(1) = -Plate_Space - valC
                            Else
                                             By(i) = By(1) - valBtoBSpace * (i - 1)
                            End If
               Next i
        Else
               k = 2
               For i = 1 To valBoltEA / 2
                              BX(i) = valGap + valA
                              If i = 1 Then
                                             By(1) = -Plate_Space - valC
                              Else
                                             By(i) = By(1) - valBtoBSpace * (k - 1)
                                             k = k + 1
                              End If
               Next i
               k = 2
               For i = valBoltEA / 2 + 1 To valBoltEA
                              BX(i) = valGap + valA + valE
                              If i = valBoltEA / 2 + 1 Then
                                             By(i) = -Plate_Space - valC
                              Else
                                             By(i) = By(valBoltEA / 2 + 1) - valBtoBSpace * (k - 1)
                                             k = k + 1
                              End If
               Next i
        End If
    Case "Type D"
        
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        If valBoltType_Flag = "I" Then
               Plate_Depth = valA + valB
               Plate_Space = (valBdepth - Plate_Depth) / 2
               For i = 1 To valBoltEA
                              By(i) = -Plate_Space - valA
                              If i = 1 Then
                                             BX(1) = valCwt / 2 + valGap + valC
                              Else
                                             BX(2) = BX(1) + valBtoBSpace * (i - 1)
                              End If
               Next i
        Else
               Plate_Depth = valA + valB + valE
               Plate_Space = (valBdepth - Plate_Depth) / 2
               k = 2
               For i = 1 To valBoltEA / 2
                              By(i) = -Plate_Space - valA
                              If i = 1 Then
                                             BX(1) = valCwt / 2 + valGap + valC
                              Else
                                             BX(i) = BX(1) + valBtoBSpace * (i - 1)
                                             k = k + 1
                              End If
               Next i
               k = 2
               For i = valBoltEA / 2 + 1 To valBoltEA
                              By(i) = -Plate_Space - valA - valE
                              If i = valBoltEA / 2 + 1 Then
                                             BX(i) = valCwt / 2 + valGap + valC
                              Else
                                             BX(i) = BX(valbotea / 2 + 1) + valBtoBSpace * (k - 1)
                                             k = k + 1
                              End If
               Next i
        End If
End Select

ReDim Nx(1 To valBoltEA, 1 To 6) As Single
ReDim Ny(1 To valBoltEA, 1 To 6) As Single

For i = 1 To valBoltEA
    For j = 1 To 6
        Nx(i, j) = (BX(i) + (valNutDia / 2) * Sin((Seta(j) / 180) * 3.14))
        Ny(i, j) = (By(i) + (valNutDia / 2) * Cos((Seta(j) / 180) * 3.14))
    Next j
Next i

If valRL_Flag = "Type01" Then RL_Flag = "R"
If valRL_Flag = "Type02" Then RL_Flag = "L"

Open valPathName For Append As #1
    For i = 1 To valBoltEA

        Print #1, "plc_area"
        For j = 1 To 6
               If valDir_Flag = "X" Then
                              Select Case RL_Flag
                                             Case "R"
                                                            Print #1, "vert" & j & " = " & Format(Nx(i, j), "0.000") & ", " & _
                                                                                               Format(valPt + valBwt / 2, "0.000") & "," & _
                                                                                               Format(Ny(i, j), "0.000") & ","
                                             Case "L"
                                                            Print #1, "vert" & j & " = " & Format(-Nx(i, j), "0.000") & ", " & _
                                                                                               Format(valPt + valBwt / 2 + valNutHei, "0.000") & "," & _
                                                                                               Format(Ny(i, j), "0.000") & ","
                              End Select
               Else
                              Select Case RL_Flag
                                             Case "R"
                                                            Print #1, "vert" & j & " = " & Format(valBwt / 2 + valPt, "0.000") & ", " & _
                                                                                               Format(-Nx(i, j), "0.000") & "," & _
                                                                                               Format(Ny(i, j), "0.000") & ","
                                             Case "L"
                                                            Print #1, "vert" & j & " = " & Format(valBwt / 2 + valPt + valNutHei, "0.000") & ", " & _
                                                                                               Format(Nx(i, j), "0.000") & "," & _
                                                                                               Format(Ny(i, j), "0.000") & ","
                              End Select
               End If
          Next j
        Print #1, "class = " & gstr_SCClass & ", " & _
                  "grade = """ & gstr_Grade & """, " & _
                  "material = """ & gstr_Material & """, " & _
                  "name = ""HTB_" & CStr(valBoltName) & """" & ", "
        Print #1, "thickness = " & valNutHei & ";"
    Next i
    
Close #1

End Sub

Public Sub SCData_Call(ByVal valJobName As String, ByVal valCode As String, ByVal valFormCode As String, _
                                                valColumn As String, valBeam As String, valSubBeam As String, valVectorDir As String, _
                                                  valModule As String)
Dim xSQL As String
Dim reData As ADODB.Recordset
Dim reData1 As ADODB.Recordset

xSQL = "select member_name,type,HTB_Name,HTB_Num,HTB_SNum,HTB_Space,Plate_Thk,Stiff_Thk," & _
              "Gap,A,B,C,D,E,Unit from SC_Connection "

If valModule = "Module01" Then
               xSQL = xSQL & " where Member_Name = '" & valBeam & "' "
               xSQL = xSQL & "and Job = '" & valJobName & "' "
               xSQL = xSQL & "and Code = '" & valCode & "'"
Else
               xSQL = xSQL & " where Member_Name = '" & valSubBeam & "' "
               xSQL = xSQL & "and Job = '" & valJobName & "' "
               xSQL = xSQL & "and Code = '" & valCode & "'"
End If
Set reData = adoConnection1.Execute(xSQL)
'
'If valVectorDir = "VectorY" Then
'    gsin_Xlen = reData!Xlen
'    gsin_Ylen = reData!Ylen
'
'    gsin_Xbtob = reData!Xbtob
'    gsin_Xcls = reData!Xcls
'    gsin_Ybtob = reData!Ybtob
'    gsin_Ycls = reData!Ycls
'Else
'    gsin_Xlen = reData!Ylen
'    gsin_Ylen = reData!Xlen
'    gsin_Xbtob = reData!Ybtob
'    gsin_Xcls = reData!Ycls
'    gsin_Ybtob = reData!Xbtob
'    gsin_Ycls = reData!Xcls
'End If

If Not reData.EOF Then
    gsin_Gap = reData!Gap
    gsin_A = reData!A
    gsin_B = reData!B
    gsin_C = reData!C
    gsin_D = reData!D
    gsin_E = reData!E
    gsin_StiffThk = reData!stiff_thk
    gsin_Pthk = reData!Plate_Thk
    gsin_HTB_SNum = reData!HTB_SNum
    gsin_HTB_Space = reData!HTB_Space
    
    gin_BoltEA = reData!HTB_Num
    gstr_BoltName = reData!HTB_Name
    gstr_Type = reData!Type
    gstr_Unit = reData!unit
Else
    gsin_Gap = 0
    gsin_A = 0
    gsin_B = 0
    gsin_C = 0
    gsin_D = 0
    gsin_E = 0
    gsin_StiffThk = 0
    gsin_Pthk = 0
    gsin_HTB_SNum = 0
    gsin_HTB_Space = 0
    
    'gin_BoltEA = reData!BoltEA
    gstr_BoltName = "M16"
    gstr_Type = "I"
    gstr_Unit = "mm"
End If

reData.Close
Set reData = Nothing

xSQL = "select dia,nutdia,nuthei,unit from BoltNut "
xSQL = xSQL & "where Name = '" & gstr_BoltName & "' and unit = '" & gstr_Unit & "'"
Set reData1 = adoConnection.Execute(xSQL)
    gsin_BoltDia = reData1!dia
    gsin_NutDia = reData1!nutdia
    gsin_NutHei = reData1!nuthei
    gstr_NutUnit = reData1!unit
reData1.Close
Set reData1 = Nothing

If valColumn <> "N/A" Then
               xSQL = "select * from code_" & valFormCode & " where member_name = '" & valColumn & "'"
               
               Set retData = adoConnection.Execute(xSQL)
                   gsin_Cdepth = retData!D
                   gsin_Cwidth = retData!Bf
                   gsin_Cwt = retData!Tw
                   gsin_Cft = retData!Tf
               retData.Close
               Set retData = Nothing
End If

If valBeam <> "N/A" Then
               If valModule = "Module01" Then
                        xSQL = "select * from code_" & valCode & " where member_name = '" & valBeam & "'"
               Else
                        xSQL = "select * from code_" & valFormCode & " where member_name = '" & valBeam & "'"
               End If
               
               Set retData = adoConnection.Execute(xSQL)
                   gsin_Bedepth = retData!D
                   gsin_Bewidth = retData!Bf
                   gsin_Bewt = retData!Tw
                   gsin_Beft = retData!Tf
               retData.Close
               Set retData = Nothing
End If

If valSubBeam <> "N/A" Then
    xSQL = "select * from code_" & valCode & " where member_name = '" & valSubBeam & "'"
    
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

xSQL = "select Grade,Material,SC_Class from Plate_General "
xSQL = xSQL & "where job = '" & valJobName & "'"

Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    gstr_Grade = reData1!grade
    gstr_Material = reData1!material
    gstr_SCClass = reData1!sc_class
Else
    gstr_Grade = "A36"
    gstr_Material = "Steel"
    gstr_SCClass = "2"
End If
reData1.Close
Set reData1 = Nothing

End Sub



Attribute VB_Name = "mdlMCEndPlate"
Public Sub MC_PML(valPath As String, ByVal valJobName As String, ByVal valCode_Left As String, _
   ByVal valCode_Right As String, ByVal valFormCode As String, _
    valDir As String, valMCEP_Flag As String, valType As String, valUnit As String, _
    valColumn As String, valLeftBeam As String, valRightBeam As String)
    
Dim TempX1 As Single, TempX2 As Single, TempY1 As Single, TempY2 As Single
Dim TempBoltDia As Single, TempNutDia As Single, TempNutHei As Single
Dim TempBoltDia_R As Single, TempNutDia_R As Single, TempNutHei_R As Single
Dim TempL As Single, TempL2 As Single, TempW As Single, TempB As Single, TempC As Single, TempD As Single, _
    TempE As Single, TempF As Single, TempG As Single, TempH As Single, TempI As Single, TempJ As Single, _
    TempThk As Single, TempStiffThk As Single, TempSPATop As Single, TempSPABot As Single
Dim TempL_R As Single, TempL2_R As Single, TempW_R As Single, TempB_R As Single, TempC_R As Single, TempD_R As Single, _
    TempE_R As Single, TempF_R As Single, TempG_R As Single, TempH_R As Single, TempI_R As Single, TempJ_R As Single, _
    TempThk_R As Single, TempStiffThk_R As Single, TempSPATop_R As Single, TempSPABot_R As Single
Dim TempCdepth As Single, TempCwidth As Single, TempCft As Single, TempCwt As Single
Dim TempBedepth As Single, TempBewidth As Single, TempBewt As Single, TempBeft As Single
Dim TempRBedepth As Single, TempRBewidth As Single, TempRBewt As Single, TempRBeft As Single
Dim TempDepth_Flag As Boolean

    Call MCData_Call(valJobName, valCode_Left, valCode_Right, valFormCode, valColumn, _
                                    valLeftBeam, valRightBeam, valDir, valMCEP_Flag)
    
    TempCdepth = convert(gsin_Cdepth, "mm", valUnit)
    TempCwidth = convert(gsin_Cwidth, "mm", valUnit)
    TempCft = convert(gsin_Cft, "mm", valUnit)
    TempCwt = convert(gsin_Cwt, "mm", valUnit)
    
    TempBedepth = convert(gsin_Bedepth, "mm", valUnit)
    TempBewidth = convert(gsin_Bewidth, "mm", valUnit)
    TempBeft = convert(gsin_Beft, "mm", valUnit)
    TempBewt = convert(gsin_Bewt, "mm", valUnit)
    
    TempThk = convert(gsin_Pthk, gstr_Unit, valUnit)
    TempStiffThk = convert(gsin_StiffThk, gstr_Unit, valUnit)
    TempSPATop = convert(gsin_SPATop, gstr_Unit, valUnit)
    TempSPABot = convert(gsin_SPABot, gstr_Unit, valUnit)
    TempL = convert(gsin_L, gstr_Unit, valUnit)
    TempL2 = convert(gsin_L2, gstr_Unit, valUnit)
    TempW = convert(gsin_W, gstr_Unit, valUnit)
    TempB = convert(gsin_B, gstr_Unit, valUnit)
    TempC = convert(gsin_C, gstr_Unit, valUnit)
    TempD = convert(gsin_D, gstr_Unit, valUnit)
    TempE = convert(gsin_E, gstr_Unit, valUnit)
    TempF = convert(gsin_F, gstr_Unit, valUnit)
    TempG = convert(gsin_G, gstr_Unit, valUnit)
    TempH = convert(gsin_H, gstr_Unit, valUnit)
    TempI = convert(gsin_I, gstr_Unit, valUnit)
    TempJ = convert(gsin_J, gstr_Unit, valUnit)
    
    TempBoltDia = convert(gsin_BoltDia, gstr_NutUnit, valUnit)
    TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
    TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)
    
    TempRBedepth = convert(gsin_rBedepth, "mm", valUnit)
    TempRBewidth = convert(gsin_rBewidth, "mm", valUnit)
    TempRBeft = convert(gsin_rBeft, "mm", valUnit)
    TempRBewt = convert(gsin_rBewt, "mm", valUnit)
    
    TempThk_R = convert(gsin_rPthk, gstr_rUnit, valUnit)
    TempStiffThk_R = convert(gsin_rStiffThk, gstr_rUnit, valUnit)
    TempSPATop_R = convert(gsin_rSPATop, gstr_rUnit, valUnit)
    TempSPABot_R = convert(gsin_rSPABot, gstr_rUnit, valUnit)
    TempL_R = convert(gsin_rL, gstr_rUnit, valUnit)
    TempL2_R = convert(gsin_rL2, gstr_rUnit, valUnit)
    TempW_R = convert(gsin_rW, gstr_rUnit, valUnit)
    TempB_R = convert(gsin_rB, gstr_rUnit, valUnit)
    TempC_R = convert(gsin_rC, gstr_rUnit, valUnit)
    TempD_R = convert(gsin_rD, gstr_rUnit, valUnit)
    TempE_R = convert(gsin_rE, gstr_rUnit, valUnit)
    TempF_R = convert(gsin_rF, gstr_rUnit, valUnit)
    TempG_R = convert(gsin_rG, gstr_rUnit, valUnit)
    TempH_R = convert(gsin_rH, gstr_rUnit, valUnit)
    TempI_R = convert(gsin_rI, gstr_rUnit, valUnit)
    TempJ_R = convert(gsin_rJ, gstr_rUnit, valUnit)
    
    TempBoltDia_R = convert(gsin_rBoltDia, gstr_rNutUnit, valUnit)
    TempNutDia_R = convert(gsin_rNutDia, gstr_rNutUnit, valUnit)
    TempNutHei_R = convert(gsin_rNutHei, gstr_rNutUnit, valUnit)
    
    Select Case valDir
        Case "X"
            Select Case valMCEP_Flag
                Case "Module01"
                    Select Case valType
                        Case "Type01"
                            Call MC_Module01_XZ_Type01(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R)
                                
                            Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                        Case "Type02"
                            Call MC_Module01_XZ_Type02(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot)
    
                            Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                TempL, TempL2, TempW, TempB, TempC, TempD, _
                                TempE, TempF, TempG, TempH, TempI, TempJ, _
                                TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                        Case "Type03"
                                    If gsin_Bedepth > gsin_rBedepth Then
                                    
                                                   If gsin_Bedepth - gsin_rBedepth <= 150 Then
                                                            TempDepth_Flag = False
                                                   Else
                                                            TempDepth_Flag = True
                                                   End If
                                                   Call MC_Module01_XZ_Type03(valPath, _
                                                   TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                   TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R, _
                                                   TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot, TempDepth_Flag)
                                                   
                                                   Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                                   TempL, TempL2, TempW, TempB, TempC, TempD, _
                                                   TempE, TempF, TempG, TempH, TempI, TempJ, _
                                                   TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                                   
                                                   Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                                   TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                                   TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                                   TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                                    Else
                                                   MsgBox "Your Selection is Wrong... "
                                                   Exit Sub
                                    End If
                            Case "Type04"
                                    If gsin_Bedepth < gsin_rBedepth Then
                                    
                                                   If gsin_rBedepth - gsin_Bedepth <= 150 Then
                                                                  TempDepth_Flag = False
                                                   Else
                                                                  TempDepth_Flag = True
                                                   End If
                                                   
                                                   Call MC_Module01_XZ_Type04(valPath, _
                                                   TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                   TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot, _
                                                   TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R, _
                                                   TempDepth_Flag)
                                                   
                                                   Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                                   TempL, TempL2, TempW, TempB, TempC, TempD, _
                                                   TempE, TempF, TempG, TempH, TempI, TempJ, _
                                                   TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                                   
                                                   Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                                   TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                                   TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                                   TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                                    Else
                                                   MsgBox "Your Selection is Wrong... "
                                                   Exit Sub
                                    End If
                        Case "Type05"
                            Call MC_Module01_XZ_Type05(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot)
    
                                Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                                
                    End Select
                Case "Module02"
                    Select Case valType
                        Case "Type01"
                            Call MC_Module02_XZ_Type01(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R)
                                
                            Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                        Case "Type02"
                            Call MC_Module02_XZ_Type02(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk)
                                
                            Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                TempL, TempL2, TempW, TempB, TempC, TempD, _
                                TempE, TempF, TempG, TempH, TempI, TempJ, _
                                TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                        Case "Type03"
                              If gsin_Bedepth > gsin_rBedepth Then

                                    If gsin_Bedepth - gsin_rBedepth <= 150 Then
                                            TempDepth_Flag = False
                                    Else
                                            TempDepth_Flag = True
                                    End If
                                    Call MC_Module02_XZ_Type03(valPath, _
                                        TempCdepth, TempCwidth, TempCwt, TempCft, _
                                        TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                        TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, _
                                        TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                        TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R, TempDepth_Flag)
                                    
                                    Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                        TempL, TempL2, TempW, TempB, TempC, TempD, _
                                        TempE, TempF, TempG, TempH, TempI, TempJ, _
                                        TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                        
                                    Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                      TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                      TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                      TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                                Else
                                             MsgBox "Your Selection is Wrong... "
                                             Exit Sub
                              End If
                        Case "Type04"
                              If gsin_Bedepth < gsin_rBedepth Then
                                    If gsin_rBedepth - gsin_Bedepth <= 150 Then
                                          TempDepth_Flag = False
                                    Else
                                          TempDepth_Flag = True
                                    End If
                                          Call MC_Module02_XZ_Type04(valPath, _
                                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                                TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R, _
                                                TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                                TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, TempDepth_Flag)
                                          
                                          Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                                TempL, TempL2, TempW, TempB, TempC, TempD, _
                                                TempE, TempF, TempG, TempH, TempI, TempJ, _
                                                TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                          
                                          Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                                TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                                TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                                TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                              Else
                                           MsgBox "Your Selection is Wrong... "
                                           Exit Sub
                              End If
                    Case "Type05"
                            Call MC_Module02_XZ_Type05(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, _
                                TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R)
    
                                Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "X")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "X")
                                
                    End Select
            End Select
        Case "Y"
            Select Case valMCEP_Flag
                Case "Module01"
                    Select Case valType
                        Case "Type01"
                            Call MC_Module01_YZ_Type01(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R)
                                
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                                
                        Case "Type02"
                            Call MC_Module01_YZ_Type02(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot)
    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                TempL, TempL2, TempW, TempB, TempC, TempD, _
                                TempE, TempF, TempG, TempH, TempI, TempJ, _
                                TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                
                        Case "Type03"
                              If gsin_Bedepth > gsin_rBedepth Then
                                             If gsin_Bedepth - gsin_rBedepth <= 150 Then
                                                    TempDepth_Flag = False
                                             Else
                                                    TempDepth_Flag = True
                                             End If
                                Call MC_Module01_YZ_Type03(valPath, _
                                    TempCdepth, TempCwidth, TempCwt, TempCft, _
                                    TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R, _
                                    TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot, TempDepth_Flag)
    
                                Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                        TempL, TempL2, TempW, TempB, TempC, TempD, _
                                        TempE, TempF, TempG, TempH, TempI, TempJ, _
                                        TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                              
                               Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                       TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                       TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                       TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                            Else
                                MsgBox "Your Selection is Wrong... "
                                Exit Sub
                            End If
                        Case "Type04"
                            If gsin_Bedepth < gsin_rBedepth Then
                                             If gsin_rBedepth - gsin_Bedepth <= 150 Then
                                                            TempDepth_Flag = False
                                             Else
                                                            TempDepth_Flag = True
                                             End If
                                Call MC_Module01_YZ_Type04(valPath, _
                                    TempCdepth, TempCwidth, TempCwt, TempCft, _
                                    TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot, _
                                    TempL_R, TempW_R, TempThk_R, TempStiffThk_R, TempSPATop_R, TempSPABot_R, TempDepth_Flag)
    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                            Else
                                MsgBox "Your Selection is Wrong... "
                                Exit Sub
                            End If
                        Case "Type05"
                            Call MC_Module01_YZ_Type05(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempL, TempW, TempThk, TempStiffThk, TempSPATop, TempSPABot)
    
                            Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                    End Select
                Case "Module02"
                    Select Case valType
                        Case "Type01"
                            Call MC_Module02_YZ_Type01(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R)
                                
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                                
                        Case "Type02"
                            Call MC_Module02_YZ_Type02(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk)
    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                TempL, TempL2, TempW, TempB, TempC, TempD, _
                                TempE, TempF, TempG, TempH, TempI, TempJ, _
                                TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                
                        Case "Type03"
                            If gsin_Bedepth > gsin_rBedepth Then
                                    If gsin_Bedepth - gsin_rBedepth <= 150 Then
                                         TempDepth_Flag = False
                                    Else
                                         TempDepth_Flag = True
                                    End If
                                             
                                Call MC_Module02_YZ_Type03(valPath, _
                                    TempCdepth, TempCwidth, TempCwt, TempCft, _
                                    TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                    TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, _
                                    TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                    TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R, TempDepth_Flag)
                                    
                                          Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                               TempL, TempL2, TempW, TempB, TempC, TempD, _
                                               TempE, TempF, TempG, TempH, TempI, TempJ, _
                                               TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                               
                                         Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                             TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                             TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                             TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                            Else
                                MsgBox "Your Selection is Wrong... "
                                Exit Sub
                            End If
                        Case "Type04"
                            If gsin_Bedepth < gsin_rBedepth Then
                                    If gsin_rBedepth - gsin_Bedepth <= 150 Then
                                          TempDepth_Flag = False
                                    Else
                                          TempDepth_Flag = True
                                    End If
                                             
                                Call MC_Module02_YZ_Type04(valPath, _
                                    TempCdepth, TempCwidth, TempCwt, TempCft, _
                                    TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                    TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R, _
                                    TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                    TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, TempDepth_Flag)
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                            Else
                                MsgBox "Your Selection is Wrong... "
                                Exit Sub
                            End If
                        Case "Type05"
                            Call MC_Module02_YZ_Type05(valPath, _
                                TempCdepth, TempCwidth, TempCwt, TempCft, _
                                TempBedepth, TempBewidth, TempBewt, TempBeft, _
                                TempL, TempW, TempThk, TempL2, TempSPABot, TempStiffThk, _
                                TempRBedepth, TempRBewidth, TempRBewt, TempRBeft, _
                                TempL_R, TempW_R, TempThk_R, TempL2_R, TempSPABot_R, TempStiffThk_R)
                                
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk, _
                                    TempL, TempL2, TempW, TempB, TempC, TempD, _
                                    TempE, TempF, TempG, TempH, TempI, TempJ, _
                                    TempBoltDia, TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, "L", "Y")
                                    
                              Call PrintNut_Moment(valPath, TempCdepth, TempThk_R, _
                                  TempL_R, TempL2_R, TempW_R, TempB_R, TempC_R, TempD_R, _
                                  TempE_R, TempF_R, TempG_R, TempH_R, TempI_R, TempJ_R, _
                                  TempBoltDia_R, TempNutDia_R, TempNutHei_R, gstr_rBoltName, gstr_rType, "R", "Y")
                                
                    End Select
            End Select
    End Select
End Sub

Public Sub MC_Module01_XZ_Type01(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px, Py, Pzt,"
    Print #1, "vert2 = Px, Py, Pzb,"
    Print #1, "vert3 = Px, -Py, Pzb,"
    Print #1, "vert4 = Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, 0,"
    Print #1, "vert2 = -spx1, spy1, 0,"
    Print #1, "vert3 = -spx1, spy2, 0,"
    Print #1, "vert4 = spx1, spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -STt,"
    Print #1, "vert2 = -spx1, -spy1, -STt,"
    Print #1, "vert3 = -spx1, -spy2, -STt,"
    Print #1, "vert4 = spx1, -spy2, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ1,"
    Print #1, "vert2 = -spx1, spy1, -SPZ1,"
    Print #1, "vert3 = -spx1, spy2, -SPZ1,"
    Print #1, "vert4 = spx1, spy2, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ2,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
    Print #1, "vert4 = spx1, -spy2, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module01_XZ_Type02(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY LEFT BEAM EXISTING"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px, Py, Pzt,"
    Print #1, "vert2 = -Px, Py, Pzb,"
    Print #1, "vert3 = -Px, -Py, Pzb,"
    Print #1, "vert4 = -Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, 0,"
    Print #1, "vert2 = -spx1, spy1, 0,"
    Print #1, "vert3 = -spx1, spy2, 0,"
    Print #1, "vert4 = spx1, spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -STt,"
    Print #1, "vert2 = -spx1, -spy1, -STt,"
    Print #1, "vert3 = -spx1, -spy2, -STt,"
    Print #1, "vert4 = spx1, -spy2, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ1,"
    Print #1, "vert2 = -spx1, spy1, -SPZ1,"
    Print #1, "vert3 = -spx1, spy2, -SPZ1,"
    Print #1, "vert4 = spx1, spy2, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ2,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
    Print #1, "vert4 = spx1, -spy2, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1

End Sub

Public Sub MC_Module01_XZ_Type03(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valSTt1 As Single, valSPA_Top1 As Single, valSPA_Bot1 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT-LAGER BEAM, RIGHT-SMALLER BEAM"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    'lager Beam"
    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP1 = " & CStr(valSPA_Top1) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT1  = " & CStr(valSPA_Bot1) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign Px1 = cd/2, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = SPA_TOP1, var_type = ""float"";"
    Print #1, "assign Pzb1 = -(EPL1-SPA_TOP1), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'lager Beam"
    Print #1, "assign SPZ3 = EPL1-SPA_TOP1-SPA_BOT1-STt1, var_type = ""float"";"
    Print #1, "assign SPZ4 = EPL1-SPA_TOP1-SPA_BOT1, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px, Py, Pzt,"
    Print #1, "vert2 = Px, Py, Pzb,"
    Print #1, "vert3 = Px, -Py, Pzb,"
    Print #1, "vert4 = Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py1, Pzt1,"
    Print #1, "vert2 = -Px1, Py1, Pzb1,"
    Print #1, "vert3 = -Px1, -Py1, Pzb1,"
    Print #1, "vert4 = -Px1, -Py1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    'Lager Beam - STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, 0,"
    Print #1, "vert2 = -spx1, spy1, 0,"
    Print #1, "vert3 = -spx1, spy2, 0,"
    Print #1, "vert4 = spx1, spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -STt1,"
    Print #1, "vert2 = -spx1, -spy1, -STt1,"
    Print #1, "vert3 = -spx1, -spy2, -STt1,"
    Print #1, "vert4 = spx1, -spy2, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ3,"
    Print #1, "vert2 = -spx1, spy1, -SPZ3,"
    Print #1, "vert3 = -spx1, spy2, -SPZ3,"
    Print #1, "vert4 = spx1, spy2, -SPZ3,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ4,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ4,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ4,"
    Print #1, "vert4 = spx1, -spy2, -SPZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    If valDepth_Flag = True Then
               'Smaller Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
               Print #1, "plc_area"
               Print #1, "vert1 = spx1, spy1, -SPZ1,"
               Print #1, "vert2 = -spx1, spy1, -SPZ1,"
               Print #1, "vert3 = -spx1, spy2, -SPZ1,"
               Print #1, "vert4 = spx1, spy2, -SPZ1,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
               'Smaller Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
               Print #1, "plc_area"
               Print #1, "vert1 = spx1, -spy1, -SPZ2,"
               Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
               Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
               Print #1, "vert4 = spx1, -spy2, -SPZ2,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
    End If
Close #1
End Sub

Public Sub MC_Module01_XZ_Type04(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valSTt1 As Single, valSPA_Top1 As Single, valSPA_Bot1 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT-SMALLER BEAM, RIGHT-LAGER BEAM"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    'lager Beam"
    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP1 = " & CStr(valSPA_Top1) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT1  = " & CStr(valSPA_Bot1) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign Px1 = cd/2+EPT1, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = SPA_TOP1, var_type = ""float"";"
    Print #1, "assign Pzb1 = -(EPL1-SPA_TOP1), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'lager Beam"
    Print #1, "assign SPZ3 = EPL1-SPA_TOP1-SPA_BOT1-STt1, var_type = ""float"";"
    Print #1, "assign SPZ4 = EPL1-SPA_TOP1-SPA_BOT1, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px, Py, Pzt,"
    Print #1, "vert2 = -Px, Py, Pzb,"
    Print #1, "vert3 = -Px, -Py, Pzb,"
    Print #1, "vert4 = -Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py1, Pzt1,"
    Print #1, "vert2 = Px1, Py1, Pzb1,"
    Print #1, "vert3 = Px1, -Py1, Pzb1,"
    Print #1, "vert4 = Px1, -Py1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    'Lager Beam - STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, 0,"
    Print #1, "vert2 = -spx1, spy1, 0,"
    Print #1, "vert3 = -spx1, spy2, 0,"
    Print #1, "vert4 = spx1, spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -STt1,"
    Print #1, "vert2 = -spx1, -spy1, -STt1,"
    Print #1, "vert3 = -spx1, -spy2, -STt1,"
    Print #1, "vert4 = spx1, -spy2, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ3,"
    Print #1, "vert2 = -spx1, spy1, -SPZ3,"
    Print #1, "vert3 = -spx1, spy2, -SPZ3,"
    Print #1, "vert4 = spx1, spy2, -SPZ3,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ4,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ4,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ4,"
    Print #1, "vert4 = spx1, -spy2, -SPZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

   If valDepth_Flag = True Then
               'Smaller Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
               Print #1, "plc_area"
               Print #1, "vert1 = spx1, spy1, -SPZ1,"
               Print #1, "vert2 = -spx1, spy1, -SPZ1,"
               Print #1, "vert3 = -spx1, spy2, -SPZ1,"
               Print #1, "vert4 = spx1, spy2, -SPZ1,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
               'Smaller Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
               Print #1, "plc_area"
               Print #1, "vert1 = spx1, -spy1, -SPZ2,"
               Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
               Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
               Print #1, "vert4 = spx1, -spy2, -SPZ2,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
    End If
Close #1
End Sub

Public Sub MC_Module01_XZ_Type05(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT,RIGHT BEAM SIZE EQUAL"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"


    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px, Py, Pzt,"
    Print #1, "vert2 = -Px, Py, Pzb,"
    Print #1, "vert3 = -Px, -Py, Pzb,"
    Print #1, "vert4 = -Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = Px+EPT, Py, Pzt,"
    Print #1, "vert2 = Px+EPT, Py, Pzb,"
    Print #1, "vert3 = Px+EPT, -Py, Pzb,"
    Print #1, "vert4 = Px+EPT, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, 0,"
    Print #1, "vert2 = -spx1, spy1, 0,"
    Print #1, "vert3 = -spx1, spy2, 0,"
    Print #1, "vert4 = spx1, spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -STt,"
    Print #1, "vert2 = -spx1, -spy1, -STt,"
    Print #1, "vert3 = -spx1, -spy2, -STt,"
    Print #1, "vert4 = spx1, -spy2, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ1,"
    Print #1, "vert2 = -spx1, spy1, -SPZ1,"
    Print #1, "vert3 = -spx1, spy2, -SPZ1,"
    Print #1, "vert4 = spx1, spy2, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ2,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
    Print #1, "vert4 = spx1, -spy2, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module02_XZ_Type01(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valEpl2 As Single, valEpl3 As Single, valSTt As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXSITING"
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

    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign BW = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"
    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPL3 = " & CStr(valEpl3) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign SPA_BOT = EPL-EPL3, var_type = ""float"";"

    'Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = 0, var_type = ""float"";"
    Print #1, "assign Pzb = -EPL, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = SPA_BOT+STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = SPA_BOT, var_type = ""float"";"

    Print #1, "assign BRAX1 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX2 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX3 = PX+1.75*EPL2-EPT, var_type = ""float"";"
    Print #1, "assign BRAX4 = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRAY1 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY2 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY3 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY4 = BW/2, var_type = ""float"";"

    Print #1, "assign BRAZ1 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ2 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ3 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ4 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRAX5 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX6 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX7 = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRAY5 = STt/2, var_type = ""float"";"
    Print #1, "assign BRAY6 = STt/2, var_type = ""float"";"
    Print #1, "assign BRAY7 = STt/2, var_type = ""float"";"

    Print #1, "assign BRAZ5 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ6 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ7 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y1 = -BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y2 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y3 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y4 = -BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y5 = BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y6 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y7 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y8 = BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z1 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z2 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z3 = BRA_STIFF_Z1+(BD/2)-BFT, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z4 = BRA_STIFF_Z2+(BD/2)-BFT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px, Py, Pzt,"
    Print #1, "vert2 = Px, Py, Pzb,"
    Print #1, "vert3 = Px, -Py, Pzb,"
    Print #1, "vert4 = Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -STt,"
    Print #1, "vert2 = -spx1, spy1, -STt,"
    Print #1, "vert3 = -spx1, spy2, -STt,"
    Print #1, "vert4 = spx1, spy2, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, 0,"
    Print #1, "vert2 = -spx1, -spy1, 0,"
    Print #1, "vert3 = -spx1, -spy2, 0,"
    Print #1, "vert4 = spx1, -spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ1,"
    Print #1, "vert2 = -spx1, spy1, -SPZ1,"
    Print #1, "vert3 = -spx1, spy2, -SPZ1,"
    Print #1, "vert4 = spx1, spy2, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ2,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
    Print #1, "vert4 = spx1, -spy2, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    'BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAx1, BRAy1, BRAZ1,"
    Print #1, "vert2 = BRAx2, BRAy2, BRAZ2,"
    Print #1, "vert3 = BRAx3, BRAy3, BRAZ3,"
    Print #1, "vert4 = BRAx4, BRAy4, BRAZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAx5, BRAY5, BRAZ4,"
    Print #1, "vert2 = BRAx6, BRAY6, BRAZ5,"
    Print #1, "vert3 = BRAx7, BRAY7, BRAZ6,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X+STT, BRA_STIFF_Y1, BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_X+STT, BRA_STIFF_Y2, BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_X+STT, BRA_STIFF_Y3, BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_X+STT, BRA_STIFF_Y4, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X, BRA_STIFF_Y5, BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_X, BRA_STIFF_Y6, BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_X, BRA_STIFF_Y7, BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_X, BRA_STIFF_Y8, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module02_XZ_Type02(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valEpl2 As Single, valEpl3 As Single, valSTt As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY LEFT BEAM EXSITING"
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

    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign BW = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"
    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPL3 = " & CStr(valEpl3) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign SPA_BOT = EPL-EPL3, var_type = ""float"";"

    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = 0, var_type = ""float"";"
    Print #1, "assign Pzb = -EPL, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = SPA_BOT+STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = SPA_BOT, var_type = ""float"";"

    Print #1, "assign BRAX1 = PX, var_type = ""float"";"
    Print #1, "assign BRAX2 = PX, var_type = ""float"";"
    Print #1, "assign BRAX3 = PX+1.75*EPL2, var_type = ""float"";"
    Print #1, "assign BRAX4 = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRAY1 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY2 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY3 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY4 = -(BW/2), var_type = ""float"";"

    Print #1, "assign BRAZ1 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ2 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ3 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ4 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRAX5 = PX, var_type = ""float"";"
    Print #1, "assign BRAX6 = PX, var_type = ""float"";"
    Print #1, "assign BRAX7 = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRAY5 = -STt/2, var_type = ""float"";"
    Print #1, "assign BRAY6 = -STt/2, var_type = ""float"";"
    Print #1, "assign BRAY7 = -STt/2, var_type = ""float"";"

    Print #1, "assign BRAZ5 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ6 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ7 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y1 = -BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y2 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y3 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y4 = -BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y5 = BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y6 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y7 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y8 = BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z1 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z2 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z3 = BRA_STIFF_Z1+(BD/2)-BFT, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z4 = BRA_STIFF_Z2+(BD/2)-BFT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px, Py, Pzt,"
    Print #1, "vert2 = -Px, Py, Pzb,"
    Print #1, "vert3 = -Px, -Py, Pzb,"
    Print #1, "vert4 = -Px, -Py, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -STt,"
    Print #1, "vert2 = -spx1, spy1, -STt,"
    Print #1, "vert3 = -spx1, spy2, -STt,"
    Print #1, "vert4 = spx1, spy2, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, 0,"
    Print #1, "vert2 = -spx1, -spy1, 0,"
    Print #1, "vert3 = -spx1, -spy2, 0,"
    Print #1, "vert4 = spx1, -spy2, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, spy1, -SPZ1,"
    Print #1, "vert2 = -spx1, spy1, -SPZ1,"
    Print #1, "vert3 = -spx1, spy2, -SPZ1,"
    Print #1, "vert4 = spx1, spy2, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx1, -spy1, -SPZ2,"
    Print #1, "vert2 = -spx1, -spy1, -SPZ2,"
    Print #1, "vert3 = -spx1, -spy2, -SPZ2,"
    Print #1, "vert4 = spx1, -spy2, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    'BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx1, BRAy1, BRAZ1,"
    Print #1, "vert2 = -BRAx2, BRAy2, BRAZ2,"
    Print #1, "vert3 = -BRAx3, BRAy3, BRAZ3,"
    Print #1, "vert4 = -BRAx4, BRAy4, BRAZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx5, BRAY5, BRAZ4,"
    Print #1, "vert2 = -BRAx6, BRAY6, BRAZ5,"
    Print #1, "vert3 = -BRAx7, BRAY7, BRAZ6,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRA_STIFF_X, BRA_STIFF_Y1, BRA_STIFF_Z1,"
    Print #1, "vert2 = -BRA_STIFF_X, BRA_STIFF_Y2, BRA_STIFF_Z2,"
    Print #1, "vert3 = -BRA_STIFF_X, BRA_STIFF_Y3, BRA_STIFF_Z3,"
    Print #1, "vert4 = -BRA_STIFF_X, BRA_STIFF_Y4, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = -(BRA_STIFF_X+STT), BRA_STIFF_Y5, BRA_STIFF_Z1,"
    Print #1, "vert2 = -(BRA_STIFF_X+STT), BRA_STIFF_Y6, BRA_STIFF_Z2,"
    Print #1, "vert3 = -(BRA_STIFF_X+STT), BRA_STIFF_Y7, BRA_STIFF_Z3,"
    Print #1, "vert4 = -(BRA_STIFF_X+STT), BRA_STIFF_Y8, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module02_XZ_Type03(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single, valDepth_Flag As Boolean)
    
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LAGER-LEFT BEAM,SMALLER-RIGHT BEAM"
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

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign Bft1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign Bft2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 =" & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1-EPL13, var_type = ""float"";"

    Print #1, "assign Px1 = cd/2+EPT1, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1+STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1+1.75*EPL12, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY11 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY12 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY13 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY14 = -(BW1/2), var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY15 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY16 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY17 = -STt1/2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2-EPL23, var_type = ""float"";"

    Print #1, "assign Px2 = cd/2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2/2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2+STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY21 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY22 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY23 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY24 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY25 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY26 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY27 = STt2/2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = ""float"";"


    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -Px1, Py1, Pzt1,"
    Print #1, "vert2 = -Px1, Py1, Pzb1,"
    Print #1, "vert3 = -Px1, -Py1, Pzb1,"
    Print #1, "vert4 = -Px1, -Py1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = Px2, Py2, Pzt2,"
    Print #1, "vert2 = Px2, Py2, Pzb2,"
    Print #1, "vert3 = Px2, -Py2, Pzb2,"
    Print #1, "vert4 = Px2, -Py2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -STt1,"
    Print #1, "vert2 = -spx11, spy11, -STt1,"
    Print #1, "vert3 = -spx11, spy12, -STt1,"
    Print #1, "vert4 = spx11, spy12, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, 0,"
    Print #1, "vert2 = -spx11, -spy11, 0,"
    Print #1, "vert3 = -spx11, -spy12, 0,"
    Print #1, "vert4 = spx11, -spy12, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -SPZ11,"
    Print #1, "vert2 = -spx11, spy11, -SPZ11,"
    Print #1, "vert3 = -spx11, spy12, -SPZ11,"
    Print #1, "vert4 = spx11, spy12, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, -SPZ12,"
    Print #1, "vert2 = -spx11, -spy11, -SPZ12,"
    Print #1, "vert3 = -spx11, -spy12, -SPZ12,"
    Print #1, "vert4 = spx11, -spy12, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx11, BRAy11, BRAZ11,"
    Print #1, "vert2 = -BRAx12, BRAy12, BRAZ12,"
    Print #1, "vert3 = -BRAx13, BRAy13, BRAZ13,"
    Print #1, "vert4 = -BRAx14, BRAy14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx15, BRAY15, BRAZ14,"
    Print #1, "vert2 = -BRAx16, BRAY16, BRAZ15,"
    Print #1, "vert3 = -BRAx17, BRAY17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRA_STIFF_X1, BRA_STIFF_Y11, BRA_STIFF_Z11,"
    Print #1, "vert2 = -BRA_STIFF_X1, BRA_STIFF_Y12, BRA_STIFF_Z12,"
    Print #1, "vert3 = -BRA_STIFF_X1, BRA_STIFF_Y13, BRA_STIFF_Z13,"
    Print #1, "vert4 = -BRA_STIFF_X1, BRA_STIFF_Y14, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = -(BRA_STIFF_X1+STT1), BRA_STIFF_Y15, BRA_STIFF_Z11,"
    Print #1, "vert2 = -(BRA_STIFF_X1+STT1), BRA_STIFF_Y16, BRA_STIFF_Z12,"
    Print #1, "vert3 = -(BRA_STIFF_X1+STT1), BRA_STIFF_Y17, BRA_STIFF_Z13,"
    Print #1, "vert4 = -(BRA_STIFF_X1+STT1), BRA_STIFF_Y18, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAx21, BRAy21, BRAZ21,"
    Print #1, "vert2 = BRAx22, BRAy22, BRAZ22,"
    Print #1, "vert3 = BRAx23, BRAy23, BRAZ23,"
    Print #1, "vert4 = BRAx24, BRAy24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAx25, BRAY25, BRAZ24,"
    Print #1, "vert2 = BRAx26, BRAY26, BRAZ25,"
    Print #1, "vert3 = BRAx27, BRAY27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X2+STT2, BRA_STIFF_Y21, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_X2+STT2, BRA_STIFF_Y22, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_X2+STT2, BRA_STIFF_Y23, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_X2+STT2, BRA_STIFF_Y24, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X2, BRA_STIFF_Y25, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_X2, BRA_STIFF_Y26, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_X2, BRA_STIFF_Y27, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_X2, BRA_STIFF_Y28, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

If valDepth_Flag = True Then
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -SPZ21,"
    Print #1, "vert2 = -spx11, spy11, -SPZ21,"
    Print #1, "vert3 = -spx11, spy12, -SPZ21,"
    Print #1, "vert4 = spx11, spy12, -SPZ21,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, -SPZ22,"
    Print #1, "vert2 = -spx11, -spy11, -SPZ22,"
    Print #1, "vert3 = -spx11, -spy12, -SPZ22,"
    Print #1, "vert4 = spx11, -spy12, -SPZ22,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
End If
Close #1

End Sub

Public Sub MC_Module02_XZ_Type04(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXSITING"
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

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign Bft1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign Bft2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 = " & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1-EPL13, var_type = ""float"";"

    Print #1, "assign Px1 = cd/2, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1+STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRAY11 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY12 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY13 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY14 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRAY15 = STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY16 = STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY17 = STt1/2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2-EPL23, var_type = ""float"";"

    Print #1, "assign Px2 = cd/2+EPT2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2/2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2+STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2+1.75*EPL22, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRAY21 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY22 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY23 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY24 = -(BW2/2), var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRAY25 = -STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY26 = -STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY27 = -STt2/2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py1, Pzt1,"
    Print #1, "vert2 = Px1, Py1, Pzb1,"
    Print #1, "vert3 = Px1, -Py1, Pzb1,"
    Print #1, "vert4 = Px1, -Py1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = -Px2, Py2, Pzt2,"
    Print #1, "vert2 = -Px2, Py2, Pzb2,"
    Print #1, "vert3 = -Px2, -Py2, Pzb2,"
    Print #1, "vert4 = -Px2, -Py2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -STt1,"
    Print #1, "vert2 = -spx11, spy11, -STt1,"
    Print #1, "vert3 = -spx11, spy12, -STt1,"
    Print #1, "vert4 = spx11, spy12, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, 0,"
    Print #1, "vert2 = -spx11, -spy11, 0,"
    Print #1, "vert3 = -spx11, -spy12, 0,"
    Print #1, "vert4 = spx11, -spy12, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -SPZ11,"
    Print #1, "vert2 = -spx11, spy11, -SPZ11,"
    Print #1, "vert3 = -spx11, spy12, -SPZ11,"
    Print #1, "vert4 = spx11, spy12, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, -SPZ12,"
    Print #1, "vert2 = -spx11, -spy11, -SPZ12,"
    Print #1, "vert3 = -spx11, -spy12, -SPZ12,"
    Print #1, "vert4 = spx11, -spy12, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAx11, BRAy11, BRAZ11,"
    Print #1, "vert2 = BRAx12, BRAy12, BRAZ12,"
    Print #1, "vert3 = BRAx13, BRAy13, BRAZ13,"
    Print #1, "vert4 = BRAx14, BRAy14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAx15, BRAY15, BRAZ14,"
    Print #1, "vert2 = BRAx16, BRAY16, BRAZ15,"
    Print #1, "vert3 = BRAx17, BRAY17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X1+STT1, BRA_STIFF_Y11, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_X1+STT1, BRA_STIFF_Y12, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_X1+STT1, BRA_STIFF_Y13, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_X1+STT1, BRA_STIFF_Y14, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X1, BRA_STIFF_Y15, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_X1, BRA_STIFF_Y16, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_X1, BRA_STIFF_Y17, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_X1, BRA_STIFF_Y18, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx21, BRAy21, BRAZ21,"
    Print #1, "vert2 = -BRAx22, BRAy22, BRAZ22,"
    Print #1, "vert3 = -BRAx23, BRAy23, BRAZ23,"
    Print #1, "vert4 = -BRAx24, BRAy24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx25, BRAY25, BRAZ24,"
    Print #1, "vert2 = -BRAx26, BRAY26, BRAZ25,"
    Print #1, "vert3 = -BRAx27, BRAY27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRA_STIFF_X2, BRA_STIFF_Y21, BRA_STIFF_Z21,"
    Print #1, "vert2 = -BRA_STIFF_X2, BRA_STIFF_Y22, BRA_STIFF_Z22,"
    Print #1, "vert3 = -BRA_STIFF_X2, BRA_STIFF_Y23, BRA_STIFF_Z23,"
    Print #1, "vert4 = -BRA_STIFF_X2, BRA_STIFF_Y24, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y25, BRA_STIFF_Z21,"
    Print #1, "vert2 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y26, BRA_STIFF_Z22,"
    Print #1, "vert3 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y27, BRA_STIFF_Z23,"
    Print #1, "vert4 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y28, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

If valDepth_Flag = True Then

    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -SPZ21,"
    Print #1, "vert2 = -spx11, spy11, -SPZ21,"
    Print #1, "vert3 = -spx11, spy12, -SPZ21,"
    Print #1, "vert4 = spx11, spy12, -SPZ21,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, -SPZ22,"
    Print #1, "vert2 = -spx11, -spy11, -SPZ22,"
    Print #1, "vert3 = -spx11, -spy12, -SPZ22,"
    Print #1, "vert4 = spx11, -spy12, -SPZ22,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
End If
Close #1

End Sub

Public Sub MC_Module02_XZ_Type05(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'RIGHT,LEFT BEAM EQUAL"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""float"";"
    Print #1, "assign endy=%%point_y, var_type=""float"";"
    Print #1, "assign endz=%%point_z, var_type=""float"";"

    Print #1, "origin local = endx, endy, endz;"

    '------------------ Data Input Start ------------------------"

    Print #1, "assign cd = " & CStr(valCd) & ", var_type = ""float"";"
    Print #1, "assign Cw = " & CStr(valCw) & ", var_type = ""float"";"
    Print #1, "assign cwt = " & CStr(valCwt) & ", var_type = ""float"";"
    Print #1, "assign cft = " & CStr(valCft) & ", var_type = ""float"";"

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign BFT1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign BFT2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 = " & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1 - EPL13, var_type = ""float"";"

    Print #1, "assign PX1 = cd / 2 , var_type = ""float"";"
    Print #1, "assign Py1 = EPW1 / 2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign spl = cd - (cft * 2), var_type = ""float"";"
    Print #1, "assign spw = (Cw - cwt) / 2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl / 2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt / 2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt / 2) + spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1 + STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1+ EPT1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1+ EPT1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1 + 1.75 * EPL12 - EPT1, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1 + 1.75 * EPL12 - EPT1, var_type = ""float"";"

    Print #1, "assign BRAY11 = BW1 / 2, var_type = ""float"";"
    Print #1, "assign BRAY12 = -(BW1 / 2), var_type = ""float"";"
    Print #1, "assign BRAY13 = -(BW1 / 2), var_type = ""float"";"
    Print #1, "assign BRAY14 = BW1 / 2, var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12 + EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12 + EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1+ EPT1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1+ EPT1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1 + 1.75 * EPL12 - EPT1, var_type = ""float"";"

    Print #1, "assign BRAY15 = STt1 / 2, var_type = ""float"";"
    Print #1, "assign BRAY16 = STt1 / 2, var_type = ""float"";"
    Print #1, "assign BRAY17 = STt1 / 2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12 + EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12 + EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1 + 1.75 * EPL12 - EPT1, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1 / 2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1 / 2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1 + BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1 + BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11 + (BD1 / 2) - BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12 + (BD1 / 2) - BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2 - EPL23, var_type = ""float"";"

    Print #1, "assign PX2 = cd / 2 + EPT2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2 / 2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2 + STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2 , var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2 , var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2 + 1.75 * EPL22, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2 + 1.75 * EPL22, var_type = ""float"";"

    Print #1, "assign BRAY21 = -(BW2 / 2), var_type = ""float"";"
    Print #1, "assign BRAY22 = BW2 / 2, var_type = ""float"";"
    Print #1, "assign BRAY23 = BW2 / 2, var_type = ""float"";"
    Print #1, "assign BRAY24 = -(BW2 / 2), var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22 + EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22 + EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2 , var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2 , var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2 + 1.75 * EPL22, var_type = ""float"";"

    Print #1, "assign BRAY25 = -STt2 / 2, var_type = ""float"";"
    Print #1, "assign BRAY26 = -STt2 / 2, var_type = ""float"";"
    Print #1, "assign BRAY27 = -STt2 / 2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22 + EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22 + EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2 + 1.75 * EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2 / 2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2 / 2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2 / 2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2 + BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2 + BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21 + (BD2 / 2) - BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22 + (BD2 / 2) - BFT2, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Px1, Py1, Pzt1,"
    Print #1, "vert2 = Px1, Py1, Pzb1,"
    Print #1, "vert3 = Px1, -Py1, Pzb1,"
    Print #1, "vert4 = Px1, -Py1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = -Px2, Py2, Pzt2,"
    Print #1, "vert2 = -Px2, Py2, Pzb2,"
    Print #1, "vert3 = -Px2, -Py2, Pzb2,"
    Print #1, "vert4 = -Px2, -Py2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -STt1,"
    Print #1, "vert2 = -spx11, spy11, -STt1,"
    Print #1, "vert3 = -spx11, spy12, -STt1,"
    Print #1, "vert4 = spx11, spy12, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, 0,"
    Print #1, "vert2 = -spx11, -spy11, 0,"
    Print #1, "vert3 = -spx11, -spy12, 0,"
    Print #1, "vert4 = spx11, -spy12, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, spy11, -SPZ11,"
    Print #1, "vert2 = -spx11, spy11, -SPZ11,"
    Print #1, "vert3 = -spx11, spy12, -SPZ11,"
    Print #1, "vert4 = spx11, spy12, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = spx11, -spy11, -SPZ12,"
    Print #1, "vert2 = -spx11, -spy11, -SPZ12,"
    Print #1, "vert3 = -spx11, -spy12, -SPZ12,"
    Print #1, "vert4 = spx11, -spy12, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAx11, BRAy11, BRAZ11,"
    Print #1, "vert2 = BRAx12, BRAy12, BRAZ12,"
    Print #1, "vert3 = BRAx13, BRAy13, BRAZ13,"
    Print #1, "vert4 = BRAx14, BRAy14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAx15, BRAY15, BRAZ14,"
    Print #1, "vert2 = BRAx16, BRAY16, BRAZ15,"
    Print #1, "vert3 = BRAx17, BRAY17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X1+STT1, BRA_STIFF_Y11, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_X1+STT1, BRA_STIFF_Y12, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_X1+STT1, BRA_STIFF_Y13, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_X1+STT1, BRA_STIFF_Y14, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_X1, BRA_STIFF_Y15, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_X1, BRA_STIFF_Y16, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_X1, BRA_STIFF_Y17, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_X1, BRA_STIFF_Y18, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx21, BRAy21, BRAZ21,"
    Print #1, "vert2 = -BRAx22, BRAy22, BRAZ22,"
    Print #1, "vert3 = -BRAx23, BRAy23, BRAZ23,"
    Print #1, "vert4 = -BRAx24, BRAy24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRAx25, BRAY25, BRAZ24,"
    Print #1, "vert2 = -BRAx26, BRAY26, BRAZ25,"
    Print #1, "vert3 = -BRAx27, BRAY27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -BRA_STIFF_X2, BRA_STIFF_Y21, BRA_STIFF_Z21,"
    Print #1, "vert2 = -BRA_STIFF_X2, BRA_STIFF_Y22, BRA_STIFF_Z22,"
    Print #1, "vert3 = -BRA_STIFF_X2, BRA_STIFF_Y23, BRA_STIFF_Z23,"
    Print #1, "vert4 = -BRA_STIFF_X2, BRA_STIFF_Y24, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y25, BRA_STIFF_Z21,"
    Print #1, "vert2 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y26, BRA_STIFF_Z22,"
    Print #1, "vert3 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y27, BRA_STIFF_Z23,"
    Print #1, "vert4 = -(BRA_STIFF_X2+STT2), BRA_STIFF_Y28, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
'    Print #1, "plc_area"
'    Print #1, "vert1 = spx11, spy11, -SPZ21,"
'    Print #1, "vert2 = -spx11, spy11, -SPZ21,"
'    Print #1, "vert3 = -spx11, spy12, -SPZ21,"
'    Print #1, "vert4 = spx11, spy12, -SPZ21,"
'    Print #1, "class = " & gstr_MCClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""SP_" & CStr(valSTt2) & """" & ", "
'    Print #1, "thickness = STt2;"
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
'    Print #1, "plc_area"
'    Print #1, "vert1 = spx11, -spy11, -SPZ22,"
'    Print #1, "vert2 = -spx11, -spy11, -SPZ22,"
'    Print #1, "vert3 = -spx11, -spy12, -SPZ22,"
'    Print #1, "vert4 = spx11, -spy12, -SPZ22,"
'    Print #1, "class = " & gstr_MCClass & ", " & _
'              "grade = """ & gstr_Grade & """, " & _
'              "material = """ & gstr_Material & """, " & _
'              "name = ""SP_" & CStr(valSTt2) & """" & ", "
'    Print #1, "thickness = STt2;"
Close #1
End Sub

Public Sub MC_Module01_YZ_Type01(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"

    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"

    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px, Pzt,"
    Print #1, "vert2 = Py, -Px, Pzb,"
    Print #1, "vert3 = -Py, -Px, Pzb,"
    Print #1, "vert4 = -Py, -Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, 0,"
    Print #1, "vert2 = spy2, spx1, 0,"
    Print #1, "vert3 = spy2, -spx1, 0,"
    Print #1, "vert4 = spy1, -spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -STt,"
    Print #1, "vert2 = -spy2, spx1, -STt,"
    Print #1, "vert3 = -spy2, -spx1, -STt,"
    Print #1, "vert4 = -spy1, -spx1, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -SPZ1,"
    Print #1, "vert2 = spy2, spx1, -SPZ1,"
    Print #1, "vert3 = spy2, -spx1, -SPZ1,"
    Print #1, "vert4 = spy1, -spx1, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -SPZ2,"
    Print #1, "vert2 = -spy2, spx1, -SPZ2,"
    Print #1, "vert3 = -spy2, -spx1, -SPZ2,"
    Print #1, "vert4 = -spy1, -spx1, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1

End Sub

Public Sub MC_Module01_YZ_Type02(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY LEFT BEAM EXISTING"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px, Pzt,"
    Print #1, "vert2 = Py, Px, Pzb,"
    Print #1, "vert3 = -Py, Px, Pzb,"
    Print #1, "vert4 = -Py, Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy2, -spx1, 0,"
    Print #1, "vert2 = spy1, -spx1, 0,"
    Print #1, "vert3 = spy1, spx1, 0,"
    Print #1, "vert4 = spy2, spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy2, -spx1, -STt,"
    Print #1, "vert2 = -spy1, -spx1, -STt,"
    Print #1, "vert3 = -spy1, spx1, -STt,"
    Print #1, "vert4 = -spy2, spx1, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy2, -spx1, -SPZ1,"
    Print #1, "vert2 = spy1, -spx1, -SPZ1,"
    Print #1, "vert3 = spy1, spx1, -SPZ1,"
    Print #1, "vert4 = spy2, spx1, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy2, -spx1, -SPZ2,"
    Print #1, "vert2 = -spy1, -spx1, -SPZ2,"
    Print #1, "vert3 = -spy1, spx1, -SPZ2,"
    Print #1, "vert4 = -spy2, spx1, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1

End Sub

Public Sub MC_Module01_YZ_Type03(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valSTt1 As Single, valSPA_Top1 As Single, valSPA_Bot1 As Single, valDepth_Flag)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT-LAGER BEAM, RIGHT-SMALLER BEAM"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    'lager Beam"
    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP1 = " & CStr(valSPA_Top1) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT1  = " & CStr(valSPA_Bot1) & ", var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign Px1 = cd/2, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = SPA_TOP1, var_type = ""float"";"
    Print #1, "assign Pzb1 = -(EPL1-SPA_TOP1), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'lager Beam"
    Print #1, "assign SPZ3 = EPL1-SPA_TOP1-SPA_BOT1-STt1, var_type = ""float"";"
    Print #1, "assign SPZ4 = EPL1-SPA_TOP1-SPA_BOT1, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px, Pzt,"
    Print #1, "vert2 = Py, -Px, Pzb,"
    Print #1, "vert3 = -Py, -Px, Pzb,"
    Print #1, "vert4 = -Py, -Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py1, Px1, Pzt1,"
    Print #1, "vert2 = Py1, Px1, Pzb1,"
    Print #1, "vert3 = -Py1, Px1, Pzb1,"
    Print #1, "vert4 = -Py1, Px1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    'Lager Beam - STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, 0,"
    Print #1, "vert2 = spy2, spx1, 0,"
    Print #1, "vert3 = spy2, -spx1, 0,"
    Print #1, "vert4 = spy1, -spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -STt1,"
    Print #1, "vert2 = -spy2, spx1, -STt1,"
    Print #1, "vert3 = -spy2, -spx1, -STt1,"
    Print #1, "vert4 = -spy1, -spx1, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -SPZ3,"
    Print #1, "vert2 = spy2, spx1, -SPZ3,"
    Print #1, "vert3 = spy2, -spx1, -SPZ3,"
    Print #1, "vert4 = spy1, -spx1, -SPZ3,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -SPZ4,"
    Print #1, "vert2 = -spy2, spx1, -SPZ4,"
    Print #1, "vert3 = -spy2, -spx1, -SPZ4,"
    Print #1, "vert4 = -spy1, -spx1, -SPZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    If valDepth_Flag = True Then

               'Smaller Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
               Print #1, "plc_area"
               Print #1, "vert1 = spy1, spx1, -SPZ1,"
               Print #1, "vert2 = spy2, spx1, -SPZ1,"
               Print #1, "vert3 = spy2, -spx1, -SPZ1,"
               Print #1, "vert4 = spy1, -spx1, -SPZ1,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
               'Smaller Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
               Print #1, "plc_area"
               Print #1, "vert1 = -spy1, spx1, -SPZ2,"
               Print #1, "vert2 = -spy2, spx1, -SPZ2,"
               Print #1, "vert3 = -spy2, -spx1, -SPZ2,"
               Print #1, "vert4 = -spy1, -spx1, -SPZ2,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
    End If
Close #1

End Sub

Public Sub MC_Module01_YZ_Type04(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valSTt1 As Single, valSPA_Top1 As Single, valSPA_Bot1 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT-SMALLER BEAM, RIGHT-LAGER BEAM"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"

    'lager Beam"
    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP1 = " & CStr(valSPA_Top1) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT1  = " & CStr(valSPA_Bot1) & "0.095, var_type = ""float"";"

    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign Px1 = cd/2+EPT1, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = SPA_TOP1, var_type = ""float"";"
    Print #1, "assign Pzb1 = -(EPL1-SPA_TOP1), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"

    'lager Beam"
    Print #1, "assign SPZ3 = EPL1-SPA_TOP1-SPA_BOT1-STt1, var_type = ""float"";"
    Print #1, "assign SPZ4 = EPL1-SPA_TOP1-SPA_BOT1, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px, Pzt,"
    Print #1, "vert2 = Py, Px, Pzb,"
    Print #1, "vert3 = -Py, Px, Pzb,"
    Print #1, "vert4 = -Py, Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py1, -Px1, Pzt1,"
    Print #1, "vert2 = Py1, -Px1, Pzb1,"
    Print #1, "vert3 = -Py1, -Px1, Pzb1,"
    Print #1, "vert4 = -Py1, -Px1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    'Lager Beam - STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, 0,"
    Print #1, "vert2 = spy2, spx1, 0,"
    Print #1, "vert3 = spy2, -spx1, 0,"
    Print #1, "vert4 = spy1, -spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -STt1,"
    Print #1, "vert2 = -spy2, spx1, -STt1,"
    Print #1, "vert3 = -spy2, -spx1, -STt1,"
    Print #1, "vert4 = -spy1, -spx1, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -SPZ3,"
    Print #1, "vert2 = spy2, spx1, -SPZ3,"
    Print #1, "vert3 = spy2, -spx1, -SPZ3,"
    Print #1, "vert4 = spy1, -spx1, -SPZ3,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'Lager Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -SPZ4,"
    Print #1, "vert2 = -spy2, spx1, -SPZ4,"
    Print #1, "vert3 = -spy2, -spx1, -SPZ4,"
    Print #1, "vert4 = -spy1, -spx1, -SPZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    If valDepth_Flag = True Then

               'Smaller Beam - STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
               Print #1, "plc_area"
               Print #1, "vert1 = spy1, spx1, -SPZ1,"
               Print #1, "vert2 = spy2, spx1, -SPZ1,"
               Print #1, "vert3 = spy2, -spx1, -SPZ1,"
               Print #1, "vert4 = spy1, -spx1, -SPZ1,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
               'Smaller Beam -STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
               Print #1, "plc_area"
               Print #1, "vert1 = -spy1, spx1, -SPZ2,"
               Print #1, "vert2 = -spy2, spx1, -SPZ2,"
               Print #1, "vert3 = -spy2, -spx1, -SPZ2,"
               Print #1, "vert4 = -spy1, -spx1, -SPZ2,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt) & """" & ", "
               Print #1, "thickness = STt;"
    End If
Close #1

End Sub

Public Sub MC_Module01_YZ_Type05(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valSTt As Single, valSPA_Top As Single, valSPA_Bot As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LEFT,RIGHT BEAM SIZE EQUAL"
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

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"

    Print #1, "assign SPA_TOP = " & CStr(valSPA_Top) & ", var_type = ""float"";"
    Print #1, "assign SPA_BOT = " & CStr(valSPA_Bot) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = SPA_TOP, var_type = ""float"";"
    Print #1, "assign Pzb = -(EPL-SPA_TOP), var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = EPL-SPA_TOP-SPA_BOT-STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = EPL-SPA_TOP-SPA_BOT, var_type = ""float"";"


    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px-EPT, Pzt,"
    Print #1, "vert2 = Py, -Px-EPT, Pzb,"
    Print #1, "vert3 = -Py, -Px-EPT, Pzb,"
    Print #1, "vert4 = -Py, -Px-EPT, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px, Pzt,"
    Print #1, "vert2 = Py, Px, Pzb,"
    Print #1, "vert3 = -Py, Px, Pzb,"
    Print #1, "vert4 = -Py, Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, 0,"
    Print #1, "vert2 = spy2, spx1, 0,"
    Print #1, "vert3 = spy2, -spx1, 0,"
    Print #1, "vert4 = spy1, -spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -STt,"
    Print #1, "vert2 = -spy2, spx1, -STt,"
    Print #1, "vert3 = -spy2, -spx1, -STt,"
    Print #1, "vert4 = -spy1, -spx1, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -SPZ1,"
    Print #1, "vert2 = spy2, spx1, -SPZ1,"
    Print #1, "vert3 = spy2, -spx1, -SPZ1,"
    Print #1, "vert4 = spy1, -spx1, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -SPZ2,"
    Print #1, "vert2 = -spy2, spx1, -SPZ2,"
    Print #1, "vert3 = -spy2, -spx1, -SPZ2,"
    Print #1, "vert4 = -spy1, -spx1, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1

End Sub

Public Sub MC_Module02_YZ_Type01(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valEpl2 As Single, valEpl3 As Single, valSTt As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXSITING"
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

    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign BW = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT = " & CStr(valEpt) & ", var_type = ""float"";"
    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPL3 = " & CStr(valEpl3) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign SPA_BOT = EPL-EPL3, var_type = ""float"";"

    Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = 0, var_type = ""float"";"
    Print #1, "assign Pzb = -EPL, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = SPA_BOT+STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = SPA_BOT, var_type = ""float"";"

    Print #1, "assign BRAX1 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX2 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX3 = PX+1.75*EPL2-EPT, var_type = ""float"";"
    Print #1, "assign BRAX4 = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRAY1 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY2 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY3 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY4 = BW/2, var_type = ""float"";"

    Print #1, "assign BRAZ1 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ2 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ3 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ4 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRAX5 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX6 = PX+EPT, var_type = ""float"";"
    Print #1, "assign BRAX7 = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRAY5 = STt/2, var_type = ""float"";"
    Print #1, "assign BRAY6 = STt/2, var_type = ""float"";"
    Print #1, "assign BRAY7 = STt/2, var_type = ""float"";"

    Print #1, "assign BRAZ5 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ6 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ7 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X = PX+1.75*EPL2-EPT, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y1 = -BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y2 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y3 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y4 = -BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y5 = BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y6 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y7 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y8 = BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z1 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z2 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z3 = BRA_STIFF_Z1+(BD/2)-BFT, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z4 = BRA_STIFF_Z2+(BD/2)-BFT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, -Px, Pzt,"
    Print #1, "vert2 = Py, -Px, Pzb,"
    Print #1, "vert3 = -Py, -Px, Pzb,"
    Print #1, "vert4 = -Py, -Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -STt,"
    Print #1, "vert2 = spy2, spx1, -STt,"
    Print #1, "vert3 = spy2, -spx1, -STt,"
    Print #1, "vert4 = spy1, -spx1, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, 0,"
    Print #1, "vert2 = -spy2, spx1, 0,"
    Print #1, "vert3 = -spy2, -spx1, 0,"
    Print #1, "vert4 = -spy1, -spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy1, spx1, -SPZ1,"
    Print #1, "vert2 = spy2, spx1, -SPZ1,"
    Print #1, "vert3 = spy2, -spx1, -SPZ1,"
    Print #1, "vert4 = spy1, -spx1, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy1, spx1, -SPZ2,"
    Print #1, "vert2 = -spy2, spx1, -SPZ2,"
    Print #1, "vert3 = -spy2, -spx1, -SPZ2,"
    Print #1, "vert4 = -spy1, -spx1, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    'BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy1, -BRAx1, BRAZ1,"
    Print #1, "vert2 = BRAy2, -BRAx2, BRAZ2,"
    Print #1, "vert3 = BRAy3, -BRAx3, BRAZ3,"
    Print #1, "vert4 = BRAy4, -BRAx4, BRAZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy5, -BRAx5, BRAZ4,"
    Print #1, "vert2 = BRAy6, -BRAx6, BRAZ5,"
    Print #1, "vert3 = BRAy7, -BRAx7, BRAZ6,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y1, -(BRA_STIFF_X+STT), BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_Y2, -(BRA_STIFF_X+STT), BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_Y3, -(BRA_STIFF_X+STT), BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_Y4, -(BRA_STIFF_X+STT), BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y5, -BRA_STIFF_X, BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_Y6, -BRA_STIFF_X, BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_Y7, -BRA_STIFF_X, BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_Y8, -BRA_STIFF_X, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module02_YZ_Type02(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd As Single, valBw As Single, valBwt As Single, valBft As Single, _
    valEpl As Single, valEpw As Single, valEpt As Single, _
    valEpl2 As Single, valEpl3 As Single, valSTt As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY LEFT BEAM EXSITING"
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

    Print #1, "assign BD = " & CStr(valBd) & ", var_type = ""float"";"
    Print #1, "assign BW = " & CStr(valBw) & ", var_type = ""float"";"
    Print #1, "assign Bwt = " & CStr(valBwt) & ", var_type = ""float"";"
    Print #1, "assign Bft = " & CStr(valBft) & ", var_type = ""float"";"

    Print #1, "assign EPL = " & CStr(valEpl) & ", var_type = ""float"";"
    Print #1, "assign EPW = " & CStr(valEpw) & ", var_type = ""float"";"
    Print #1, "assign EPT =" & CStr(valEpt) & ", var_type = ""float"";"
    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPL3 = " & CStr(valEpl3) & ", var_type = ""float"";"

    Print #1, "assign STt = " & CStr(valSTt) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    Print #1, "assign SPA_BOT = EPL-EPL3, var_type = ""float"";"

    'Print #1, "assign Px = cd/2, var_type = ""float"";"
    Print #1, "assign Px = cd/2+EPT, var_type = ""float"";"
    Print #1, "assign Py = EPW/2, var_type = ""float"";"
    Print #1, "assign Pzt = 0, var_type = ""float"";"
    Print #1, "assign Pzb = -EPL, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX1 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY1 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY2 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ1 = SPA_BOT+STt, var_type = ""float"";"
    Print #1, "assign SPZ2 = SPA_BOT, var_type = ""float"";"

    Print #1, "assign BRAX1 = PX, var_type = ""float"";"
    Print #1, "assign BRAX2 = PX, var_type = ""float"";"
    Print #1, "assign BRAX3 = PX+1.75*EPL2, var_type = ""float"";"
    Print #1, "assign BRAX4 = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRAY1 = -(BW/2), var_type = ""float"";"
    Print #1, "assign BRAY2 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY3 = BW/2, var_type = ""float"";"
    Print #1, "assign BRAY4 = -(BW/2), var_type = ""float"";"

    Print #1, "assign BRAZ1 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ2 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ3 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ4 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRAX5 = PX, var_type = ""float"";"
    Print #1, "assign BRAX6 = PX, var_type = ""float"";"
    Print #1, "assign BRAX7 = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRAY5 = -STt/2, var_type = ""float"";"
    Print #1, "assign BRAY6 = -STt/2, var_type = ""float"";"
    Print #1, "assign BRAY7 = -STt/2, var_type = ""float"";"

    Print #1, "assign BRAZ5 = -SPZ2, var_type = ""float"";"
    Print #1, "assign BRAZ6 = -SPZ2+EPL2, var_type = ""float"";"
    Print #1, "assign BRAZ7 = -SPZ2+EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X = PX+1.75*EPL2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y1 = -BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y2 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y3 = -Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y4 = -BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y5 = BW/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y6 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y7 = Bwt/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y8 = BW/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z1 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z2 = -BD+Bft, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z3 = BRA_STIFF_Z1+(BD/2)-BFT, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z4 = BRA_STIFF_Z2+(BD/2)-BFT, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py, Px, Pzt,"
    Print #1, "vert2 = Py, Px, Pzb,"
    Print #1, "vert3 = -Py, Px, Pzb,"
    Print #1, "vert4 = -Py, Px, Pzt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt) & """" & ", "
    Print #1, "thickness = EPT;"

    'STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy2, -spx1, -STt,"
    Print #1, "vert2 = spy1, -spx1, -STt,"
    Print #1, "vert3 = spy1, spx1, -STt,"
    Print #1, "vert4 = spy2, spx1, -STt,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy2, -spx1, 0,"
    Print #1, "vert2 = -spy1, -spx1, 0,"
    Print #1, "vert3 = -spy1, spx1, 0,"
    Print #1, "vert4 = -spy2, spx1, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy2, -spx1, -SPZ1,"
    Print #1, "vert2 = spy1, -spx1, -SPZ1,"
    Print #1, "vert3 = spy1, spx1, -SPZ1,"
    Print #1, "vert4 = spy2, spx1, -SPZ1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
    'STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy2, -spx1, -SPZ2,"
    Print #1, "vert2 = -spy1, -spx1, -SPZ2,"
    Print #1, "vert3 = -spy1, spx1, -SPZ2,"
    Print #1, "vert4 = -spy2, spx1, -SPZ2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    'BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy1, BRAx1, BRAZ1,"
    Print #1, "vert2 = BRAy2, BRAx2, BRAZ2,"
    Print #1, "vert3 = BRAy3, BRAx3, BRAZ3,"
    Print #1, "vert4 = BRAy4, BRAx4, BRAZ4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy5, BRAx5, BRAZ4,"
    Print #1, "vert2 = BRAy6, BRAx6, BRAZ5,"
    Print #1, "vert3 = BRAy7, BRAx7, BRAZ6,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y1, BRA_STIFF_X, BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_Y2, BRA_STIFF_X, BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_Y3, BRA_STIFF_X, BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_Y4, BRA_STIFF_X, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y5, BRA_STIFF_X+STT, BRA_STIFF_Z1,"
    Print #1, "vert2 = BRA_STIFF_Y6, BRA_STIFF_X+STT, BRA_STIFF_Z2,"
    Print #1, "vert3 = BRA_STIFF_Y7, BRA_STIFF_X+STT, BRA_STIFF_Z3,"
    Print #1, "vert4 = BRA_STIFF_Y8, BRA_STIFF_X+STT, BRA_STIFF_Z4,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt) & """" & ", "
    Print #1, "thickness = STt;"
Close #1
End Sub

Public Sub MC_Module02_YZ_Type03(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LAGER-LEFT BEAM,SMALLER-RIGHT BEAM"
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

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign Bft1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign Bft2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 = " & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1-EPL13, var_type = ""float"";"

    Print #1, "assign Px1 = cd/2+EPT1, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1+STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1+1.75*EPL12, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY11 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY12 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY13 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY14 = -(BW1/2), var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY15 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY16 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY17 = -STt1/2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2-EPL23, var_type = ""float"";"

    Print #1, "assign Px2 = cd/2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2/2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2+STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY21 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY22 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY23 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY24 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY25 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY26 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY27 = STt2/2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = ""float"";"


    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py1, Px1, Pzt1,"
    Print #1, "vert2 = Py1, Px1, Pzb1,"
    Print #1, "vert3 = -Py1, Px1, Pzb1,"
    Print #1, "vert4 = -Py1, Px1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py2, -Px2, Pzt2,"
    Print #1, "vert2 = Py2, -Px2, Pzb2,"
    Print #1, "vert3 = -Py2, -Px2, Pzb2,"
    Print #1, "vert4 = -Py2, -Px2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -STt1,"
    Print #1, "vert2 = spy12, spx11, -STt1,"
    Print #1, "vert3 = spy12, -spx11, -STt1,"
    Print #1, "vert4 = spy11, -spx11, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, 0,"
    Print #1, "vert2 = -spy12, spx11, 0,"
    Print #1, "vert3 = -spy12, -spx11, 0,"
    Print #1, "vert4 = -spy11, -spx11, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -SPZ11,"
    Print #1, "vert2 = spy12, spx11, -SPZ11,"
    Print #1, "vert3 = spy12, -spx11, -SPZ11,"
    Print #1, "vert4 = spy11, -spx11, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, -SPZ12,"
    Print #1, "vert2 = -spy12, spx11, -SPZ12,"
    Print #1, "vert3 = -spy12, -spx11, -SPZ12,"
    Print #1, "vert4 = -spy11, -spx11, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy11, BRAx11, BRAZ11,"
    Print #1, "vert2 = BRAy12, BRAx12, BRAZ12,"
    Print #1, "vert3 = BRAy13, BRAx13, BRAZ13,"
    Print #1, "vert4 = BRAy14, BRAx14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy15, BRAx15, BRAZ14,"
    Print #1, "vert2 = BRAy16, BRAx16, BRAZ15,"
    Print #1, "vert3 = BRAy17, BRAx17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y11, BRA_STIFF_X1, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y12, BRA_STIFF_X1, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y13, BRA_STIFF_X1,  BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y14, BRA_STIFF_X1, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y15, BRA_STIFF_X1+STT1, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y16, BRA_STIFF_X1+STT1, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y17, BRA_STIFF_X1+STT1, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y18, BRA_STIFF_X1+STT1, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy21, -BRAx21, BRAZ21,"
    Print #1, "vert2 = BRAy22, -BRAx22, BRAZ22,"
    Print #1, "vert3 = BRAy23, -BRAx23, BRAZ23,"
    Print #1, "vert4 = BRAy24, -BRAx24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy25, -BRAx25, BRAZ24,"
    Print #1, "vert2 = BRAy26, -BRAx26, BRAZ25,"
    Print #1, "vert3 = BRAy27, -BRAx27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y21, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y22, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y23, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y24, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y25, -BRA_STIFF_X2, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y26, -BRA_STIFF_X2, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y27, -BRA_STIFF_X2, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y28, -BRA_STIFF_X2, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
      
    If valDepth_Flag = True Then
               'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
               Print #1, "plc_area"
               Print #1, "vert1 = spy11, spx11, -SPZ21,"
               Print #1, "vert2 = spy12, spx11, -SPZ21,"
               Print #1, "vert3 = spy12, -spx11, -SPZ21,"
               Print #1, "vert4 = spy11, -spx11, -SPZ21,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt2) & """" & ", "
               Print #1, "thickness = STt2;"
               'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
               Print #1, "plc_area"
               Print #1, "vert1 = -spy11, spx11, -SPZ22,"
               Print #1, "vert2 = -spy12, spx11, -SPZ22,"
               Print #1, "vert3 = -spy12, -spx11, -SPZ22,"
               Print #1, "vert4 = -spy11, -spx11, -SPZ22,"
               Print #1, "class = " & gstr_MCClass & ", " & _
                         "grade = """ & gstr_Grade & """, " & _
                         "material = """ & gstr_Material & """, " & _
                         "name = ""SP_" & CStr(valSTt2) & """" & ", "
               Print #1, "thickness = STt2;"
     End If
Close #1

End Sub

Public Sub MC_Module02_YZ_Type04(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single, valDepth_Flag As Boolean)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'ONLY RIGHT BEAM EXSITING"
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

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign Bft1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 = " & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign Bft2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 = " & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1-EPL13, var_type = ""float"";"

    Print #1, "assign Px1 = cd/2, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1+STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRAY11 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY12 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY13 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY14 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1+EPT1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRAY15 = STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY16 = STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY17 = STt1/2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1+1.75*EPL12-EPT1, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2-EPL23, var_type = ""float"";"

    Print #1, "assign Px2 = cd/2+EPT2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2/2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2+STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2+1.75*EPL22, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRAY21 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY22 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY23 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY24 = -(BW2/2), var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2, var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRAY25 = -STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY26 = -STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY27 = -STt2/2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2+1.75*EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = ""float"";"

    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py1, -Px1, Pzt1,"
    Print #1, "vert2 = Py1, -Px1, Pzb1,"
    Print #1, "vert3 = -Py1, -Px1, Pzb1,"
    Print #1, "vert4 = -Py1, -Px1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py2, Px2, Pzt2,"
    Print #1, "vert2 = Py2, Px2, Pzb2,"
    Print #1, "vert3 = -Py2, Px2, Pzb2,"
    Print #1, "vert4 = -Py2, Px2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -STt1,"
    Print #1, "vert2 = spy12, spx11, -STt1,"
    Print #1, "vert3 = spy12, -spx11, -STt1,"
    Print #1, "vert4 = spy11, -spx11, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, 0,"
    Print #1, "vert2 = -spy12, spx11, 0,"
    Print #1, "vert3 = -spy12, -spx11, 0,"
    Print #1, "vert4 = -spy11, -spx11, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -SPZ11,"
    Print #1, "vert2 = spy12, spx11, -SPZ11,"
    Print #1, "vert3 = spy12, -spx11, -SPZ11,"
    Print #1, "vert4 = spy11, -spx11, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, -SPZ12,"
    Print #1, "vert2 = -spy12, spx11, -SPZ12,"
    Print #1, "vert3 = -spy12, -spx11, -SPZ12,"
    Print #1, "vert4 = -spy11, -spx11, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy11, -BRAx11, BRAZ11,"
    Print #1, "vert2 = BRAy12, -BRAx12, BRAZ12,"
    Print #1, "vert3 = BRAy13, -BRAx13, BRAZ13,"
    Print #1, "vert4 = BRAy14, -BRAx14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy15, -BRAx15, BRAZ14,"
    Print #1, "vert2 = BRAy16, -BRAx16, BRAZ15,"
    Print #1, "vert3 = BRAy17, -BRAx17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y11, -(BRA_STIFF_X1+STT1), BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y12, -(BRA_STIFF_X1+STT1), BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y13, -(BRA_STIFF_X1+STT1), BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y14, -(BRA_STIFF_X1+STT1), BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y15, -BRA_STIFF_X1, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y16, -BRA_STIFF_X1, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y17, -BRA_STIFF_X1, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y18, -BRA_STIFF_X1, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy21, BRAx21, BRAZ21,"
    Print #1, "vert2 = BRAy22, BRAx22, BRAZ22,"
    Print #1, "vert3 = BRAy23, BRAx23, BRAZ23,"
    Print #1, "vert4 = BRAy24, BRAx24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy25, BRAx25, BRAZ24,"
    Print #1, "vert2 = BRAy26, BRAx26, BRAZ25,"
    Print #1, "vert3 = BRAy27, BRAx27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y21, BRA_STIFF_X2, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y22, BRA_STIFF_X2, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y23, BRA_STIFF_X2, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y24, BRA_STIFF_X2, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y25, BRA_STIFF_X2+STT2, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y26, BRA_STIFF_X2+STT2, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y27, BRA_STIFF_X2+STT2, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y28, BRA_STIFF_X2+STT2, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

If valDepth_Flag = True Then
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -SPZ21,"
    Print #1, "vert2 = spy12, spx11, -SPZ21,"
    Print #1, "vert3 = spy12, -spx11, -SPZ21,"
    Print #1, "vert4 = spy11, -spx11, -SPZ21,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
    'SMALLER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, -SPZ22,"
    Print #1, "vert2 = -spy12, spx11, -SPZ22,"
    Print #1, "vert3 = -spy12, -spx11, -SPZ22,"
    Print #1, "vert4 = -spy11, -spx11, -SPZ22,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
End If
Close #1

End Sub

Public Sub MC_Module02_YZ_Type05(valPathName As String, _
    valCd As Single, valCw As Single, valCwt As Single, valCft As Single, _
    valBd1 As Single, valBw1 As Single, valBwt1 As Single, valBft1 As Single, _
    valEpl1 As Single, valEpw1 As Single, valEpt1 As Single, _
    valEpl12 As Single, valEpl13 As Single, valSTt1 As Single, _
    valBd2 As Single, valBw2 As Single, valBwt2 As Single, valBft2 As Single, _
    valEpl2 As Single, valEpw2 As Single, valEpt2 As Single, _
    valEpl22 As Single, valEpl23 As Single, valSTt2 As Single)
    
Open valPathName For Output As #1
    Print #1, "Default delete_log = ""yes"";"
    'LAGER-LEFT BEAM,SMALLER-RIGHT BEAM"
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

    'LAGER BEAM DATA"
    Print #1, "assign BD1 = " & CStr(valBd1) & ", var_type = ""float"";"
    Print #1, "assign BW1 = " & CStr(valBw1) & ", var_type = ""float"";"
    Print #1, "assign Bwt1 = " & CStr(valBwt1) & ", var_type = ""float"";"
    Print #1, "assign Bft1 = " & CStr(valBft1) & ", var_type = ""float"";"

    Print #1, "assign EPL1 = " & CStr(valEpl1) & ", var_type = ""float"";"
    Print #1, "assign EPW1 = " & CStr(valEpw1) & ", var_type = ""float"";"
    Print #1, "assign EPT1 = " & CStr(valEpt1) & ", var_type = ""float"";"
    Print #1, "assign EPL12 = " & CStr(valEpl12) & ", var_type = ""float"";"
    Print #1, "assign EPL13 = " & CStr(valEpl13) & ", var_type = ""float"";"

    Print #1, "assign STt1 =" & CStr(valSTt1) & ", var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign BD2 = " & CStr(valBd2) & ", var_type = ""float"";"
    Print #1, "assign BW2 = " & CStr(valBw2) & ", var_type = ""float"";"
    Print #1, "assign Bwt2 = " & CStr(valBwt2) & ", var_type = ""float"";"
    Print #1, "assign Bft2 = " & CStr(valBft2) & ", var_type = ""float"";"

    Print #1, "assign EPL2 = " & CStr(valEpl2) & ", var_type = ""float"";"
    Print #1, "assign EPW2 = " & CStr(valEpw2) & ", var_type = ""float"";"
    Print #1, "assign EPT2 = " & CStr(valEpt2) & ", var_type = ""float"";"
    Print #1, "assign EPL22 = " & CStr(valEpl22) & ", var_type = ""float"";"
    Print #1, "assign EPL23 = " & CStr(valEpl23) & ", var_type = ""float"";"

    Print #1, "assign STt2 = " & CStr(valSTt2) & ", var_type = ""float"";"


    '------------------ Data Input End ------------------------"

    'LAGER BEAM DATA"
    Print #1, "assign SPA_BOT1 = EPL1-EPL13, var_type = ""float"";"

    Print #1, "assign Px1 = cd/2+EPT1, var_type = ""float"";"
    Print #1, "assign Py1 = EPW1/2, var_type = ""float"";"
    Print #1, "assign Pzt1 = 0, var_type = ""float"";"
    Print #1, "assign Pzb1 = -EPL1, var_type = ""float"";"

    Print #1, "assign SPl = cd-(cft*2), var_type = ""float"";"
    Print #1, "assign SPw = (cw-cwt)/2, var_type = ""float"";"

    Print #1, "assign SPX11 = spl/2, var_type = ""float"";"
    Print #1, "assign SPY11 = cwt/2, var_type = ""float"";"
    Print #1, "assign SPY12 = (cwt/2)+spw, var_type = ""float"";"
    Print #1, "assign SPZ11 = SPA_BOT1+STt1, var_type = ""float"";"
    Print #1, "assign SPZ12 = SPA_BOT1, var_type = ""float"";"

    Print #1, "assign BRAX11 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX12 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX13 = PX1+1.75*EPL12, var_type = ""float"";"
    Print #1, "assign BRAX14 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY11 = -(BW1/2), var_type = ""float"";"
    Print #1, "assign BRAY12 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY13 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRAY14 = -(BW1/2), var_type = ""float"";"

    Print #1, "assign BRAZ11 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ12 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ13 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ14 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRAX15 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX16 = PX1, var_type = ""float"";"
    Print #1, "assign BRAX17 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRAY15 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY16 = -STt1/2, var_type = ""float"";"
    Print #1, "assign BRAY17 = -STt1/2, var_type = ""float"";"

    Print #1, "assign BRAZ15 = -SPZ12, var_type = ""float"";"
    Print #1, "assign BRAZ16 = -SPZ12+EPL12, var_type = ""float"";"
    Print #1, "assign BRAZ17 = -SPZ12+EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X1 = PX1+1.75*EPL12, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y11 = -BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y12 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y13 = -Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y14 = -BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y15 = BW1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y16 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y17 = Bwt1/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y18 = BW1/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z11 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z12 = -BD1+Bft1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z13 = BRA_STIFF_Z11+(BD1/2)-BFT1, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z14 = BRA_STIFF_Z12+(BD1/2)-BFT1, var_type = ""float"";"

    'SMALLER BEAM DATA"
    Print #1, "assign SPA_BOT2 = EPL2-EPL23, var_type = ""float"";"

    Print #1, "assign Px2 = cd/2, var_type = ""float"";"
    Print #1, "assign Py2 = EPW2/2, var_type = ""float"";"
    Print #1, "assign Pzt2 = 0, var_type = ""float"";"
    Print #1, "assign Pzb2 = -EPL2, var_type = ""float"";"

    Print #1, "assign SPZ21 = SPA_BOT2+STt2, var_type = ""float"";"
    Print #1, "assign SPZ22 = SPA_BOT2, var_type = ""float"";"

    Print #1, "assign BRAX21 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX22 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX23 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"
    Print #1, "assign BRAX24 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY21 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRAY22 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY23 = -(BW2/2), var_type = ""float"";"
    Print #1, "assign BRAY24 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRAZ21 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ22 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ23 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ24 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRAX25 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX26 = PX2+EPT2, var_type = ""float"";"
    Print #1, "assign BRAX27 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRAY25 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY26 = STt2/2, var_type = ""float"";"
    Print #1, "assign BRAY27 = STt2/2, var_type = ""float"";"

    Print #1, "assign BRAZ25 = -SPZ22, var_type = ""float"";"
    Print #1, "assign BRAZ26 = -SPZ22+EPL22, var_type = ""float"";"
    Print #1, "assign BRAZ27 = -SPZ22+EPL22, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_X2 = PX2+1.75*EPL22-EPT2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y21 = -BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y22 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y23 = -Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y24 = -BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Y25 = BW2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y26 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y27 = Bwt2/2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Y28 = BW2/2, var_type = ""float"";"

    Print #1, "assign BRA_STIFF_Z21 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z22 = -BD2+Bft2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z23 = BRA_STIFF_Z21+(BD2/2)-BFT2, var_type = ""float"";"
    Print #1, "assign BRA_STIFF_Z24 = BRA_STIFF_Z22+(BD2/2)-BFT2, var_type = ""float"";"


    'END PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = Py1, Px1, Pzt1,"
    Print #1, "vert2 = Py1, Px1, Pzb1,"
    Print #1, "vert3 = -Py1, Px1, Pzb1,"
    Print #1, "vert4 = -Py1, Px1, Pzt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt1) & """" & ", "
    Print #1, "thickness = EPT1;"

    Print #1, "plc_area"
    Print #1, "vert1 = Py2, -Px2, Pzt2,"
    Print #1, "vert2 = Py2, -Px2, Pzb2,"
    Print #1, "vert3 = -Py2, -Px2, Pzb2,"
    Print #1, "vert4 = -Py2, -Px2, Pzt2,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""EP_" & CStr(valEpt2) & """" & ", "
    Print #1, "thickness = EPT2;"

    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -STt1,"
    Print #1, "vert2 = spy12, spx11, -STt1,"
    Print #1, "vert3 = spy12, -spx11, -STt1,"
    Print #1, "vert4 = spy11, -spx11, -STt1,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR TOP-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, 0,"
    Print #1, "vert2 = -spy12, spx11, 0,"
    Print #1, "vert3 = -spy12, -spx11, 0,"
    Print #1, "vert4 = -spy11, -spx11, 0,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-LEFT"
    Print #1, "plc_area"
    Print #1, "vert1 = spy11, spx11, -SPZ11,"
    Print #1, "vert2 = spy12, spx11, -SPZ11,"
    Print #1, "vert3 = spy12, -spx11, -SPZ11,"
    Print #1, "vert4 = spy11, -spx11, -SPZ11,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"
    'LAGER BEAM-STIFFNER PLATE MODELING FOR BOTTOM-RIGHT"
    Print #1, "plc_area"
    Print #1, "vert1 = -spy11, spx11, -SPZ12,"
    Print #1, "vert2 = -spy12, spx11, -SPZ12,"
    Print #1, "vert3 = -spy12, -spx11, -SPZ12,"
    Print #1, "vert4 = -spy11, -spx11, -SPZ12,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'LAGER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy11, BRAx11, BRAZ11,"
    Print #1, "vert2 = BRAy12, BRAx12, BRAZ12,"
    Print #1, "vert3 = BRAy13, BRAx13, BRAZ13,"
    Print #1, "vert4 = BRAy14, BRAx14, BRAZ14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy15, BRAx15, BRAZ14,"
    Print #1, "vert2 = BRAy16, BRAx16, BRAZ15,"
    Print #1, "vert3 = BRAy17, BRAx17, BRAZ16,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y11, BRA_STIFF_X1, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y12, BRA_STIFF_X1, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y13, BRA_STIFF_X1,  BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y14, BRA_STIFF_X1, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y15, BRA_STIFF_X1+STT1, BRA_STIFF_Z11,"
    Print #1, "vert2 = BRA_STIFF_Y16, BRA_STIFF_X1+STT1, BRA_STIFF_Z12,"
    Print #1, "vert3 = BRA_STIFF_Y17, BRA_STIFF_X1+STT1, BRA_STIFF_Z13,"
    Print #1, "vert4 = BRA_STIFF_Y18, BRA_STIFF_X1+STT1, BRA_STIFF_Z14,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt1) & """" & ", "
    Print #1, "thickness = STt1;"

    'SMALLER BEAM-BRACKET PLATE MODELING"
    Print #1, "plc_area"
    Print #1, "vert1 = BRAy21, -BRAx21, BRAZ21,"
    Print #1, "vert2 = BRAy22, -BRAx22, BRAZ22,"
    Print #1, "vert3 = BRAy23, -BRAx23, BRAZ23,"
    Print #1, "vert4 = BRAy24, -BRAx24, BRAZ24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRAy25, -BRAx25, BRAZ24,"
    Print #1, "vert2 = BRAy26, -BRAx26, BRAZ25,"
    Print #1, "vert3 = BRAy27, -BRAx27, BRAZ26,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y21, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y22, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y23, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y24, -(BRA_STIFF_X2+STT2), BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"

    Print #1, "plc_area"
    Print #1, "vert1 = BRA_STIFF_Y25, -BRA_STIFF_X2, BRA_STIFF_Z21,"
    Print #1, "vert2 = BRA_STIFF_Y26, -BRA_STIFF_X2, BRA_STIFF_Z22,"
    Print #1, "vert3 = BRA_STIFF_Y27, -BRA_STIFF_X2, BRA_STIFF_Z23,"
    Print #1, "vert4 = BRA_STIFF_Y28, -BRA_STIFF_X2, BRA_STIFF_Z24,"
    Print #1, "class = " & gstr_MCClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = ""SP_" & CStr(valSTt2) & """" & ", "
    Print #1, "thickness = STt2;"
Close #1

End Sub

Public Sub MCData_Call(ByVal valJobName As String, ByVal valCode_Left As String, _
                                                ByVal valCode_Right As String, ByVal valFormCode As String, _
                                                valColumn As String, valLeftBeam As String, valRightBeam As String, _
                                                valVectorDir As String, valModlue As String)
Dim xSQL As String
Dim reData As ADODB.Recordset
Dim reData1 As ADODB.Recordset

If valLeftBeam <> "N/A" Then
    If valModlue = "Module01" Then
               xSQL = "select member_name,type,HTB_Name,HTB_Num,Plate_Thk,Stiff_Thk," & _
                           "L,L2,W,B,C,D,E,F,G,H,I,J,Unit from MC_Connection "
               xSQL = xSQL & "where Member_Name = '" & valLeftBeam & "' "
               xSQL = xSQL & "and job = '" & valJobName & "' "
               xSQL = xSQL & "and code = '" & valCode_Left & "' "
               xSQL = xSQL & "and (type = 'A1' or type = 'A2' or type = 'A3' or type = 'A4')"
    Else
               xSQL = "select member_name,type,HTB_Name,HTB_Num,Plate_Thk,Stiff_Thk," & _
                           "L,L2,W,B,C,D,E,F,G,H,I,J,Unit from MC_Connection "
               xSQL = xSQL & "where Member_Name = '" & valLeftBeam & "'"
               xSQL = xSQL & "and job = '" & valJobName & "' "
               xSQL = xSQL & "and code = '" & valCode_Left & "' "
               xSQL = xSQL & "and (type = 'B1' or type = 'B2' or type = 'B3' or type = 'B4' or type = 'B5')"
    End If
    
    Set reData = adoConnection1.Execute(xSQL)
    
    gsin_L = reData!L
    gsin_L2 = reData!L2
    gsin_W = reData!W
    gsin_B = reData!B
    gsin_C = reData!C
    gsin_D = reData!D
    gsin_E = reData!E
    gsin_F = reData!f
    gsin_G = reData!G
    gsin_H = reData!H
    gsin_I = reData!i
    gsin_J = reData!j
    gsin_StiffThk = reData!stiff_thk
    gsin_Pthk = reData!Plate_Thk
    gstr_Type = reData!Type
    gstr_Unit = reData!unit
    gstr_BoltName = reData!HTB_Name
    
    reData.Close
    Set reData = Nothing
    
    Select Case gstr_Type
        Case "A1"
            gsin_SPATop = gsin_C + gsin_D
            gsin_SPABot = gsin_F + gsin_G
        Case "A2"
            gsin_SPATop = gsin_C + gsin_D
            gsin_SPABot = gsin_G + gsin_H
        Case "A3"
            gsin_SPATop = gsin_C + gsin_D
            gsin_SPABot = gsin_G
        Case "A4"
            gsin_SPATop = gsin_C + gsin_D
            gsin_SPABot = gsin_I
        Case "B1"
            gsin_SPATop = 0
            gsin_SPABot = gsin_E
        Case "B2"
            gsin_SPATop = 0
            gsin_SPABot = gsin_F
        Case "B3"
            gsin_SPATop = 0
            gsin_SPABot = gsin_H
        Case "B4"
            gsin_SPATop = 0
            gsin_SPABot = gsin_I
        Case "B5"
            gsin_SPATop = 0
            gsin_SPABot = gsin_J
    End Select
    
    xSQL = "select dia,nutdia,nuthei,unit from BoltNut "
    xSQL = xSQL & "where Name = '" & gstr_BoltName & "' and unit = '" & gstr_Unit & "'"
    Set reData1 = adoConnection.Execute(xSQL)
        gsin_BoltDia = reData1!dia
        gsin_NutDia = reData1!nutdia
        gsin_NutHei = reData1!nuthei
        gstr_NutUnit = reData1!unit
    reData1.Close
    Set reData1 = Nothing

    xSQL = "select * from code_" & valCode_Left & "  where member_name = '" & valLeftBeam & "'"
    
    Set retData = adoConnection.Execute(xSQL)
        gsin_Bedepth = retData!D
        gsin_Bewidth = retData!Bf
        gsin_Bewt = retData!Tw
        gsin_Beft = retData!Tf
    retData.Close
End If

If valRightBeam <> "N/A" Then
        If valModlue = "Module01" Then
               xSQL = "select member_name,type,HTB_Name,HTB_Num,Plate_Thk,Stiff_Thk," & _
                           "L,L2,W,B,C,D,E,F,G,H,I,J,Unit from MC_Connection "
               xSQL = xSQL & "where Member_Name = '" & valRightBeam & "' "
               xSQL = xSQL & "and job = '" & valJobName & "' "
               xSQL = xSQL & "and code = '" & valCode_Right & "' "
               xSQL = xSQL & "and (type = 'A1' or type = 'A2' or type = 'A3' or type = 'A4')"
    Else
               xSQL = "select member_name,type,HTB_Name,HTB_Num,Plate_Thk,Stiff_Thk," & _
                           "L,L2,W,B,C,D,E,F,G,H,I,J,Unit from MC_Connection "
               xSQL = xSQL & "where Member_Name = '" & valRightBeam & "'"
               xSQL = xSQL & "and job = '" & valJobName & "' "
               xSQL = xSQL & "and code = '" & valCode_Right & "' "
               xSQL = xSQL & "and (type = 'B1' or type = 'B2' or type = 'B3' or type = 'B4' or type = 'B5')"
    End If

    Set reData = adoConnection1.Execute(xSQL)
    
    gsin_rL = reData!L
    gsin_rL2 = reData!L2
    gsin_rW = reData!W
    gsin_rB = reData!B
    gsin_rC = reData!C
    gsin_rD = reData!D
    gsin_rE = reData!E
    gsin_rF = reData!f
    gsin_rG = reData!G
    gsin_rH = reData!H
    gsin_rI = reData!i
    gsin_rJ = reData!j
    gsin_rStiffThk = reData!stiff_thk
    gsin_rPthk = reData!Plate_Thk
    gstr_rType = reData!Type
    gstr_rUnit = reData!unit
    gstr_rBoltName = reData!HTB_Name
    
    reData.Close
    Set reData = Nothing
    
    Select Case gstr_rType
        Case "A1"
            gsin_rSPATop = gsin_rC + gsin_rD
            gsin_rSPABot = gsin_rF + gsin_rG
        Case "A2"
            gsin_rSPATop = gsin_rC + gsin_rD
            gsin_rSPABot = gsin_rG + gsin_rH
        Case "A3"
            gsin_rSPATop = gsin_rC + gsin_rD
            gsin_rSPABot = gsin_rG
        Case "A4"
            gsin_rSPATop = gsin_rC + gsin_rD
            gsin_rSPABot = gsin_rI
        Case "B1"
            gsin_rSPATop = 0
            gsin_rSPABot = gsin_rE
        Case "B2"
            gsin_rSPATop = 0
            gsin_rSPABot = gsin_rF
        Case "B3"
            gsin_rSPATop = 0
            gsin_rSPABot = gsin_rH
        Case "B4"
            gsin_rSPATop = 0
            gsin_rSPABot = gsin_rI
        Case "B5"
            gsin_rSPATop = 0
            gsin_rSPABot = gsin_rJ
    End Select
    
    xSQL = "select dia,nutdia,nuthei,unit from BoltNut "
    xSQL = xSQL & "where Name = '" & gstr_rBoltName & "' and unit = '" & gstr_rUnit & "'"
    Set reData1 = adoConnection.Execute(xSQL)
        gsin_rBoltDia = reData1!dia
        gsin_rNutDia = reData1!nutdia
        gsin_rNutHei = reData1!nuthei
        gstr_rNutUnit = reData1!unit
    reData1.Close
    Set reData1 = Nothing

    xSQL = "select * from code_" & valCode_Right & "  where member_name = '" & valRightBeam & "'"
    
    Set retData = adoConnection.Execute(xSQL)
        gsin_rBedepth = retData!D
        gsin_rBewidth = retData!Bf
        gsin_rBewt = retData!Tw
        gsin_rBeft = retData!Tf
    retData.Close
    
    Set retData = Nothing
End If


xSQL = "select * from code_" & valFormCode & "  where member_name = '" & valColumn & "'"

Set retData = adoConnection.Execute(xSQL)
    gsin_Cdepth = retData!D
    gsin_Cwidth = retData!Bf
    gsin_Cwt = retData!Tw
    gsin_Cft = retData!Tf
retData.Close

xSQL = "select Grade,Material,MC_Class from Plate_General "
xSQL = xSQL & "where job = '" & valJobName & "'"

Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    gstr_Grade = reData1!grade
    gstr_Material = reData1!material
    gstr_MCClass = reData1!mc_class
Else
    gstr_Grade = "A36"
    gstr_Material = "Steel"
    gstr_MCClass = "2"
End If
reData1.Close
Set reData1 = Nothing

End Sub

Public Sub PrintNut_Moment(valPathName As String, valCd As Single, valPt As Single, _
    valL As Single, valL2 As Single, valW As Single, valB As Single, valC As Single, valD As Single, _
    valE As Single, valF As Single, valG As Single, valH As Single, valI As Single, valJ As Single, _
    valBoltDia As Single, valNutDia As Single, valNutHei As Single, valBoltName As String, valType As String, _
    valDir_Flag As String, valAxis_Flag As String)


Dim Seta(1 To 6) As Single
Dim Nx() As Single, Ny() As Single
Dim BX() As Single, By() As Single
Dim i As Integer, j As Integer
Dim valBoltEA As Integer

For i = 1 To 6
    If i = 1 Then
        Seta(1) = 60
    Else
        Seta(i) = 60 + Seta(i - 1)
    End If
Next i

Select Case valType
    Case "A1"
        valBoltEA = 6
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = valD
        BX(2) = valB / 2
        By(2) = valD
        
        BX(3) = -valB / 2
        By(3) = -valE
        BX(4) = valB / 2
        By(4) = -valE
        
        BX(5) = -valB / 2
        By(5) = -valE * 2 - valF
        BX(6) = valB / 2
        By(6) = -valE * 2 - valF
    Case "A2"
        valBoltEA = 8
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = valD
        BX(2) = valB / 2
        By(2) = valD
        
        BX(3) = -valB / 2
        By(3) = -valE
        BX(4) = valB / 2
        By(4) = -valE
        
        BX(5) = -valB / 2
        By(5) = -(valL - valC - valD - valE - valF - valG - valH) - valE
        BX(6) = valB / 2
        By(6) = -(valL - valC - valD - valE - valF - valG - valH) - valE
        
        BX(7) = -valB / 2
        By(7) = -(valL - valC - valD - valE - valF - valG - valH) - valE - valF - valG
        BX(8) = valB / 2
        By(8) = -(valL - valC - valD - valE - valF - valG - valH) - valE - valF - valG
    Case "A3"
        valBoltEA = 12
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2 - valH
        By(1) = valD
        BX(2) = -valB / 2
        By(2) = valD
        BX(3) = valB / 2
        By(3) = valD
        BX(4) = valB / 2 + valH
        By(4) = valD
        
        BX(5) = -valB / 2 - valH
        By(5) = -valE
        BX(6) = -valB / 2
        By(6) = -valE
        BX(7) = valB / 2
        By(7) = -valE
        BX(8) = valB / 2 + valH
        By(8) = -valE
        
        BX(9) = -valB / 2 - valH
        By(9) = -(valL - valC - valD - valE - valF - valG) - valE
        BX(10) = -valB / 2
        By(10) = -(valL - valC - valD - valE - valF - valG) - valE
        BX(11) = valB / 2
        By(11) = -(valL - valC - valD - valE - valF - valG) - valE
        BX(12) = valB / 2 + valH
        By(12) = -(valL - valC - valD - valE - valF - valG) - valE
    Case "A4"
        valBoltEA = 16
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2 - valJ
        By(1) = valD
        BX(2) = -valB / 2
        By(2) = valD
        BX(3) = valB / 2
        By(3) = valD
        BX(4) = valB / 2 + valJ
        By(4) = valD
        
        BX(5) = -valB / 2 - valJ
        By(5) = -valE
        BX(6) = -valB / 2
        By(6) = -valE
        BX(7) = valB / 2
        By(7) = -valE
        BX(8) = valB / 2 + valJ
        By(8) = -valE
        
        BX(9) = -valB / 2
        By(9) = -valE - valF
        BX(10) = valB / 2
        By(10) = -valE - valF
        
        BX(11) = -valB / 2
        By(11) = -valE - valF - valG
        BX(12) = valB / 2
        By(12) = -valE - valF - valG
        
        BX(13) = -valB / 2 - valJ
        By(13) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valE - valF - valG
        BX(14) = -valB / 2
        By(14) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valE - valF - valG
        BX(15) = valB / 2
        By(15) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valE - valF - valG
        BX(16) = valB / 2 + valJ
        By(16) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valE - valF - valG
    Case "B1"
        valBoltEA = 4
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = -valC
        BX(2) = valB / 2
        By(2) = -valC
        
        BX(3) = -valB / 2
        By(3) = -(valL - valC - valD - valE) - valC
        BX(4) = valB / 2
        By(4) = -(valL - valC - valD - valE) - valC
    Case "B2"
        valBoltEA = 6
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = -valC
        BX(2) = valB / 2
        By(2) = -valC
        
        BX(3) = -valB / 2
        By(3) = -valC - valD
        BX(4) = valB / 2
        By(4) = -valC - valD
        
        BX(5) = -valB / 2
        By(5) = -(valL - valC - valD - valE - valF) - valC - valD
        BX(6) = valB / 2
        By(6) = -(valL - valC - valD - valE - valF) - valC - valD
    Case "B3"
        valBoltEA = 10
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = -valC
        BX(2) = valB / 2
        By(2) = -valC
        
        BX(3) = -valB / 2
        By(3) = -valC - valD
        BX(4) = valB / 2
        By(4) = -valC - valD
        
        BX(5) = -valB / 2
        By(5) = -valC - valD - valE
        BX(6) = valB / 2
        By(6) = -valC - valD - valE
        
        BX(7) = -valB / 2
        By(7) = -(valL - valC - valD - valE - valF - valG - valH) - valC - valD - valE
        BX(8) = valB / 2
        By(8) = -(valL - valC - valD - valE - valF - valG - valH) - valC - valD - valE
        
        BX(9) = -valB / 2
        By(9) = -(valL - valC - valD - valE - valF - valG - valH) - valC - valD - valE - valF
        BX(10) = valB / 2
        By(10) = -(valL - valC - valD - valE - valF - valG - valH) - valC - valD - valE - valF
    Case "B4"
        valBoltEA = 12
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = -valC
        BX(2) = valB / 2
        By(2) = -valC
        
        BX(3) = -valB / 2
        By(3) = -valC - valD
        BX(4) = valB / 2
        By(4) = -valC - valD
        
        BX(5) = -valB / 2
        By(5) = -valC - valD - valE
        BX(6) = valB / 2
        By(6) = -valC - valD - valE
        
        BX(7) = -valB / 2
        By(7) = -valC - valD - valE - valF
        BX(8) = valB / 2
        By(8) = -valC - valD - valE - valF
        
        BX(9) = -valB / 2
        By(9) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valC - valD - valE - valF
        BX(10) = valB / 2
        By(10) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valC - valD - valE - valF
        
        BX(11) = -valB / 2
        By(11) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valC - valD - valE - valF - valG
        BX(12) = valB / 2
        By(12) = -(valL - valC - valD - valE - valF - valG - valH - valI) - valC - valD - valE - valF - valG
    Case "B5"
        valBoltEA = 14
        ReDim BX(1 To valBoltEA) As Single
        ReDim By(1 To valBoltEA) As Single
        
        BX(1) = -valB / 2
        By(1) = -valC
        BX(2) = valB / 2
        By(2) = -valC
        
        BX(3) = -valB / 2
        By(3) = -valC - valD
        BX(4) = valB / 2
        By(4) = -valC - valD
        
        BX(5) = -valB / 2
        By(5) = -(valC + valD + valE)
        BX(6) = valB / 2
        By(6) = -(valC + valD + valE)
        
        BX(7) = -valB / 2
        By(7) = -(valC + valD + valE + valF)
        BX(8) = valB / 2
        By(8) = -(valC + valD + valE + valF)
        
        BX(9) = -valB / 2
        By(9) = -(valC + valD + valE + valF + valG)
        BX(10) = valB / 2
        By(10) = -(valC + valD + valE + valF + valG)
        
        BX(11) = -valB / 2
        By(11) = -(valL - valC - valD - valE - valF - valG - valH - valI - valJ) - valC - valD - valE - valF - valG
        BX(12) = valB / 2
        By(12) = -(valL - valC - valD - valE - valF - valG - valH - valI - valJ) - valC - valD - valE - valF - valG
        
        BX(13) = -valB / 2
        By(13) = -(valL - valC - valD - valE - valF - valG - valH - valI - valJ) - valC - valD - valE - valF - valG - valH
        BX(14) = valB / 2
        By(14) = -(valL - valC - valD - valE - valF - valG - valH - valI - valJ) - valC - valD - valE - valF - valG - valH
End Select

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
        If valAxis_Flag = "X" Then
               If valDir_Flag = "R" Then
                      For j = 1 To 6
                             Print #1, "vert" & j & " = " & Format(valCd / 2 + valPt, "0.000") & ", " & _
                                                                               Format(Nx(i, j), "0.000") & "," & _
                                                                               Format(Ny(i, j), "0.000") & ","
                      Next j
               Else
                      For j = 1 To 6
                             Print #1, "vert" & j & " = " & Format(-1 * (valCd / 2 + valPt + valNutHei), "0.000") & ", " & _
                                                                               Format(Nx(i, j), "0.000") & "," & _
                                                                               Format(Ny(i, j), "0.000") & ","
                      Next j
               End If
        Else
                If valDir_Flag = "L" Then
                      For j = 1 To 6
                             Print #1, "vert" & j & " = " & Format(Nx(i, j), "0.000") & ", " & _
                                                                               Format(valCd / 2 + valPt + valNutHei, "0.000") & "," & _
                                                                               Format(Ny(i, j), "0.000") & ","
                      Next j
               Else
                      For j = 1 To 6
                             Print #1, "vert" & j & " = " & Format(Nx(i, j), "0.000") & ", " & _
                                                                               Format(-1 * (valCd / 2 + valPt), "0.000") & "," & _
                                                                               Format(Ny(i, j), "0.000") & ","
                      Next j
               End If
        End If
        
        Print #1, "class = " & gstr_MCClass & ", " & _
                  "grade = """ & gstr_Grade & """, " & _
                  "material = """ & gstr_Material & """, " & _
                  "name = ""HTB_" & CStr(valBoltName) & """" & ", "
        Print #1, "thickness = " & valNutHei & ";"
    Next i
    
Close #1

End Sub



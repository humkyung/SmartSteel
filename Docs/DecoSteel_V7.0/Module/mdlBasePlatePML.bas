Attribute VB_Name = "mdlBasePlatePML"
Public Sub BasePlate_Hinged_PML(valPath As String, valJobName As String, valMember As String, valUnit As String, _
                                valDir As String, valNut As Integer, valDM As String, valRDN As String)
Dim TempX1 As Single, TempX2 As Single, TempY1 As Single, TempY2 As Single, TempThk As Single
Dim TempBoltDia As Single, TempXbtob As Single, TempXcls As Single, TempYbtob As Single, _
    TempYcls As Single, TempBoltEA As Integer, TempNutDia As Single, TempNutHei As Single, _
    TempName As String
    
    Call BPHData_Call(valMember, valDir, valJobName)
    
    TempX1 = convert((gsin_Xlen / 2), gstr_Unit, valUnit)
    TempX2 = convert(-(gsin_Xlen / 2), gstr_Unit, valUnit)
    TempY1 = convert((gsin_Ylen / 2), gstr_Unit, valUnit)
    TempY2 = convert(-(gsin_Ylen / 2), gstr_Unit, valUnit)
    TempThk = convert(gsin_Pthk, gstr_Unit, valUnit)
    TempName = "BP" & "_" & TempThk
    
    Call Print_Preference_OnePoint(valPath)
    Call Print_RecBox(valPath, TempX1, TempX2, TempY1, TempY2, TempThk, TempName)
    
    If valNut = 1 Then
        TempXbtob = convert(gsin_Xbtob, gstr_Unit, valUnit)
        TempXcls = convert(gsin_Xcls, gstr_Unit, valUnit)
        TempYbtob = convert(gsin_Ybtob, gstr_Unit, valUnit)
        TempYcls = convert(gsin_Ycls, gstr_Unit, valUnit)
        TempBoltEA = gin_BoltEA
        TempBoltDia = convert(gsin_BoltDia, gstr_NutUnit, valUnit)
        TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
        TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)
        Call Print_Nut(valPath, TempThk, TempXbtob, TempXcls, TempYbtob, TempYcls, TempBoltEA, TempBoltDia, _
                       TempNutDia, TempNutHei, gstr_BoltName, gstr_Type, valDir)
    End If

End Sub

Public Sub BasePlate_Fixed_PML(valPath As String, ByVal valJobName As String, ByVal valCode As String, _
                                                                        valMember As String, valBoltName As String, _
                                                                        valType As String, valUnit As String, valMemUnit As String, _
                                                                        valDir As String, valNut As Integer, valA As Single, valB As Single, _
                                                                        valC As Single, valD As Single, valF As Single, valG As Single, _
                                                                        valBPt As Single, valRBt As Single, valRBr As Single, _
                                                                        valRBe As Single, valRBh As Single)
    
Dim TempXbtob As Single, TempYbtob As Single
Dim TempCdepth As Single, TempCwidth As Single, TempCwt As Single, TempCft As Single
Dim TempLx As Single, TempLy As Single, TempLz As Single
Dim TempRBt As Single, TempRBr As Single, TempRBe As Single
Dim TempRBsw As Single, TempRBww As Single, TempRBh As Single
Dim TempBoltDia As Single, TempNutDia As Single, TempNutHei As Single
Dim TempDimF As Single, TempDimG As Single
   
    Call BPFData_Call(valMember, valBoltName, valMemUnit, valJobName, valCode)

    TempCdepth = convert(gsin_Cdepth, "mm", valUnit)
    TempCwidth = convert(gsin_Cwidth, "mm", valUnit)
    TempCwt = convert(gsin_Cwt, "mm", valUnit)
    TempCft = convert(gsin_Cft, "mm", valUnit)
    TempLz = convert(valBPt, valMemUnit, valUnit)
    
    TempRBt = convert(valRBt, valMemUnit, valUnit)
    TempRBr = convert(valRBr, valMemUnit, valUnit)
    TempRBe = convert(valRBe, valMemUnit, valUnit)
    TempRBh = convert(valRBh, valMemUnit, valUnit)
    
    TempDimF = convert(valF, valMemUnit, valUnit)
    TempDimG = convert(valG, valMemUnit, valUnit)

    Call BPF_Preference_OnePoint(valPath)
    
    If valDir = "VectorY" Then
        TempLx = convert(valA, valMemUnit, valUnit)
        TempLy = convert(valB, valMemUnit, valUnit)
        
        TempXbtob = convert(valC, valMemUnit, valUnit)
        TempYbtob = convert(valD, valMemUnit, valUnit)
        
        TempRBsw = (TempLy - TempCdepth) / 2
        TempRBww = (TempLx - TempCwidth) / 2
    
        Select Case valType
            Case "Type01"
                Call BP_Fixed_PML_Type01_Y(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh, TempDimF)
            Case "Type02"
                Call BP_Fixed_PML_Type02_Y(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh, TempDimG)
            Case "Type03"
                Call BP_Fixed_PML_Type03_Y(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh)
            Case "Type04", "Type05"
                Call BP_Fixed_PML_Type04_Y(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh)
            Case "Type06", "Type07"
                Call BP_Fixed_PML_Type06(valPath, TempLx, TempLy, TempLz)
        End Select
    Else
        TempLx = convert(valB, valMemUnit, valUnit)
        TempLy = convert(valA, valMemUnit, valUnit)
        
        TempXbtob = convert(valD, valMemUnit, valUnit)
        TempYbtob = convert(valC, valMemUnit, valUnit)
        
        TempRBsw = (TempLx - TempCdepth) / 2
        TempRBww = (TempLy - TempCwidth) / 2
        
        Select Case valType
            Case "Type01"
                Call BP_Fixed_PML_Type01_X(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh, TempDimF)
            Case "Type02"
                Call BP_Fixed_PML_Type02_X(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh, TempDimG)
            Case "Type03"
                Call BP_Fixed_PML_Type03_X(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh)
            Case "Type04", "Type05"
                Call BP_Fixed_PML_Type04_X(valPath, TempCdepth, TempCwidth, TempCwt, TempCft, _
                                                    TempLx, TempLy, TempLz, _
                                                    TempRBt, TempRBr, TempRBe, _
                                                    TempRBsw, TempRBww, TempRBh)
            Case "Type06", "Type07"
                Call BP_Fixed_PML_Type06(valPath, TempLx, TempLy, TempLz)
        End Select
    End If

    If valNut = 1 Then
        TempBoltDia = convert(gsin_BoltDia, gstr_NutUnit, valUnit)
        TempNutDia = convert(gsin_NutDia, gstr_NutUnit, valUnit)
        TempNutHei = convert(gsin_NutHei, gstr_NutUnit, valUnit)
        Call BP_Fixed_PML_Nut(valPath, _
            TempLz, TempXbtob, TempYbtob, TempBoltDia, TempNutDia, TempNutHei, valBoltName, _
            valType, valDir)
    End If

End Sub

Public Sub BPHData_Call(valMName As String, valVectorDir As String, ByVal valJobName As String)
Dim xSQL As String
Dim reData As ADODB.Recordset
Dim reData1 As ADODB.Recordset
xSQL = "select Member_Name,Xlen,Ylen,Pthk,Xbtob,Xcls,Ybtob,Ycls,BoltEA,BoltName,Type,unit from BasePlate_Hinged"
xSQL = xSQL & " where Member_Name = '" & valMName & "' "
xSQL = xSQL & "and Job = '" & valJobName & "'"

Set reData = adoConnection1.Execute(xSQL)

If valVectorDir = "VectorY" Then
    gsin_Xlen = reData!Xlen
    gsin_Ylen = reData!Ylen
    
    gsin_Xbtob = reData!Xbtob
    gsin_Xcls = reData!Xcls
    gsin_Ybtob = reData!Ybtob
    gsin_Ycls = reData!Ycls
Else
    gsin_Xlen = reData!Ylen
    gsin_Ylen = reData!Xlen
    gsin_Xbtob = reData!Ybtob
    gsin_Xcls = reData!Ycls
    gsin_Ybtob = reData!Xbtob
    gsin_Ycls = reData!Xcls
End If
gsin_Pthk = reData!Pthk
gin_BoltEA = reData!BoltEA
gstr_BoltName = reData!BoltName
gstr_Type = reData!Type
gstr_Unit = reData!unit

reData.Close
Set reData = Nothing

xSQL = "select dia,nutdia,nuthei,unit from BoltNut"
xSQL = xSQL & " where Name = '" & gstr_BoltName & "' and unit = '" & gstr_Unit & "'"
Set reData1 = adoConnection.Execute(xSQL)
    gsin_BoltDia = reData1!dia
    gsin_NutDia = reData1!nutdia
    gsin_NutHei = reData1!nuthei
    gstr_NutUnit = reData1!unit
reData1.Close
Set reData1 = Nothing

'xSQL = "select * from code_JIS where member_name = '" & valMName & "'"
'
'Set retData = adoConnection.Execute(xSQL)
'    gsin_Cdepth = retData!D
'    gsin_Cwidth = retData!Bf
'    gsin_Cwt = retData!Tw
'    gsin_Cft = retData!Tf
'retData.Close
'
'Set retData = Nothing

xSQL = "select Grade,Material,BP_Class from Plate_General "
xSQL = xSQL & "where Job = '" & valJobName & "'"

Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    gstr_Grade = reData1!grade
    gstr_Material = reData1!material
    gstr_BPClass = reData1!bp_class
Else
    gstr_Grade = "A36"
    gstr_Material = "Steel"
    gstr_BPClass = "2"
End If
reData1.Close
Set reData1 = Nothing

End Sub

Public Sub BPFData_Call(ByVal valMName As String, ByVal valBoltName As String, ByVal valUnit As String, _
                                                            ByVal valJobName As String, ByVal valCode As String)
Dim xSQL As String
Dim reData As ADODB.Recordset
Dim reData1 As ADODB.Recordset

xSQL = "select dia,nutdia,nuthei,unit from BoltNut"
xSQL = xSQL & " where Name = '" & valBoltName & "' and unit = '" & valUnit & "'"
Set reData1 = adoConnection.Execute(xSQL)
    gsin_BoltDia = reData1!dia
    gsin_NutDia = reData1!nutdia
    gsin_NutHei = reData1!nuthei
    gstr_NutUnit = reData1!unit
reData1.Close
Set reData1 = Nothing

xSQL = "select * from code_" & valCode & " where member_name = '" & valMName & "'"

Set retData = adoConnection.Execute(xSQL)
    gsin_Cdepth = retData!D
    gsin_Cwidth = retData!Bf
    gsin_Cwt = retData!Tw
    gsin_Cft = retData!Tf
retData.Close

Set retData = Nothing

xSQL = "select Grade,Material,BP_Class from Plate_General "
xSQL = xSQL & "where job = '" & valJobName & "'"

Set reData1 = adoConnection1.Execute(xSQL)
If Not reData1.EOF Then
    gstr_Grade = reData1!grade
    gstr_Material = reData1!material
    gstr_BPClass = reData1!bp_class
Else
    gstr_Grade = "A36"
    gstr_Material = "Steel"
    gstr_BPClass = "3"
End If
reData1.Close
Set reData1 = Nothing

End Sub

Public Sub BPF_Preference_OnePoint(valPathName As String)
Open valPathName For Output As #1
    Print #1, "Default delete_log = " & """yes""" & ";"
    Print #1, "origin prompt = ""Define end Point"";"
    Print #1, "assign endx=%%point_x, var_type=""Float"";"
    Print #1, "assign endy=%%point_y, var_type=""Float"";"
    Print #1, "assign endz=%%point_z, var_type=""Float"";"
    Print #1, "origin local = endx, endy, endz;"
Close #1
End Sub


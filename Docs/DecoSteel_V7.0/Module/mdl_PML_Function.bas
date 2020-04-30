Attribute VB_Name = "mdl_PML_Function"
Public Sub PML_Run(ByVal valPathName As String)
Dim pathName As String
Dim intCount        As Integer
Dim strArray()      As String
Dim pmlPathName As String
Dim pmlFileName As String
Dim i As Integer
Dim FileName As String

FileName = Dir$(App.Path & "\USTNmacroPath.ini")

If Len(FileName) > 0 Then
            Open App.Path & "\" & FileName For Input As #1
                        Input #1, pathName
            Close #1
Else
            FileName = App.Path & "\USTNmacroPath.ini"
            Open FileName For Output As #1
                        Print #1, "C:\Bentley\Workspace\standards\macros\pmlrun.bas"
            Close #1
            Open FileName For Input As #1
                        Input #1, pathName
            Close #1
End If

strArray = Split(valPathName, "\")
intCount = UBound(strArray, 1)

For i = 0 To intCount - 1
            If i = 0 Then
                        pmlPathName = strArray(i)
            Else
                        pmlPathName = pmlPathName & "\" & strArray(i)
            End If
Next i

pmlFileName = strArray(intCount)

Open pathName For Output As #1
            Print #1, "Sub Main"
            Print #1, "MbeSendCommand " & """PML """
            Print #1, "MbeSendCommand " & """MDL COMMAND MGDSHOOK,fileList_setDirectoryCmd " & pmlPathName & "\"""
            Print #1, "MbeSendCommand " & """MDL COMMAND MGDSHOOK,fileList_setFileNameCmd " & pmlFileName & """"
            Print #1, "End Sub"
Close #1


'strArray = Split(valPathName, "\")
'intCount = UBound(strArray, 1)
'
'For i = 0 To intCount - 1
'            If i = 0 Then
'                        pmlPathName = strArray(i)
'            Else
'                        pmlPathName = pmlPathName & "\" & strArray(i)
'            End If
'Next i
'
'pathName = App.Path & "\pmlpath.ini"
'pmlFileName = strArray(intCount)
'
'Open pathName For Output As #1
'      Print #1, pmlPathName & "\"
'      Print #1, pmlFileName
'Close #1


End Sub


Public Sub Print_Preference_OnePoint(valPathName As String)


Open valPathName For Output As #1

    Print #1, "Default delete_log = ""yes"";"
    Print #1, "origin prompt =" & """Made by LG E&C ->Pick Location Point""" & ";"
    Print #1, "assign endx=%%point_x, var_type=" & """float""" & ";"
    Print #1, "assign endy=%%point_y, var_type=" & """float""" & ";"
    Print #1, "assign endz=%%point_z, var_type=" & """float""" & ";"
    Print #1, "origin local = endx, endy, endz;"

Close #1

End Sub

Public Sub Print_RecBox(valPathName As String, valX1 As Single, valX2 As Single, _
                        valY1 As Single, valY2 As Single, valThk As Single, valName As String)

Open valPathName For Append As #1

    Print #1, "plc_area"
    Print #1, "vert1 = " & Format(valX1, "0.000") & ", " & Format(valY1, "0.000") & ", 0,"
    Print #1, "vert2 = " & Format(valX2, "0.000") & ", " & Format(valY1, "0.000") & ", 0,"
    Print #1, "vert3 = " & Format(valX2, "0.000") & ", " & Format(valY2, "0.000") & ", 0,"
    Print #1, "vert4 = " & Format(valX1, "0.000") & ", " & Format(valY2, "0.000") & ", 0,"
    Print #1, "class = " & gstr_BPClass & ", " & _
              "grade = """ & gstr_Grade & """, " & _
              "material = """ & gstr_Material & """, " & _
              "name = """ & valName & """, " & _
              "ov_parall = 1" & ", "
    
'    Print #1, "Class = 2, grade = " & """A36""" & ", material = " & """Steel""" & ", name = """ & _
'               valName & """, "
    Print #1, "thickness = " & valThk & ";"
Close #1
End Sub
Public Sub Print_Nut(valPathName As String, valPt As Single, valXbtob As Single, valXcls As Single, _
                     valYbtob As Single, valYcls As Single, valBoltEA As Integer, valBoltDia As Single, _
                     valNutDia As Single, valNutHei As Single, valBoltName As String, valType As String, _
                     valDir As String)


Dim Seta(1 To 6) As Single
Dim Nx() As Single, Ny() As Single
Dim BX() As Single, By() As Single
Dim i As Integer, j As Integer

For i = 1 To 6
    If i = 1 Then
        Seta(1) = 60
    Else
        Seta(i) = 60 + Seta(i - 1)
    End If
Next i

ReDim BX(1 To valBoltEA) As Single
ReDim By(1 To valBoltEA) As Single
If valDir = "VectorY" Then
      Select Case valType
          Case "I"
              BX(1) = valXbtob / 2: BX(2) = -(valXbtob / 2)
              By(1) = 0: By(2) = 0
          Case "II"
              BX(1) = valXbtob / 2: BX(2) = valXbtob / 2
              BX(3) = -(valXbtob / 2): BX(4) = -(valXbtob / 2)
              
              By(1) = valYbtob / 2: By(2) = -(valYbtob / 2)
              By(3) = valYbtob / 2: By(4) = -(valYbtob / 2)
          Case "III"
              BX(1) = valXbtob / 2: BX(2) = valXbtob / 2: BX(3) = valXbtob / 2
              BX(4) = -(valXbtob / 2): BX(5) = -(valXbtob / 2): BX(6) = -(valXbtob / 2)
              
              By(1) = valYbtob: By(2) = 0: By(3) = -1 * valYbtob
              By(4) = valYbtob: By(5) = 0: By(6) = -1 * valYbtob
          Case "IV"
              BX(1) = valXbtob / 2: BX(2) = valXbtob / 2: BX(3) = valXbtob / 2: BX(4) = valXbtob / 2
              BX(5) = -(valXbtob / 2): BX(6) = -(valXbtob / 2): BX(7) = -(valXbtob / 2): BX(8) = -(valXbtob / 2)
              
              By(1) = valYbtob + valYbtob / 2
              By(2) = (valYbtob / 2)
              By(3) = -(valYbtob / 2)
              By(4) = -(valYbtob + valYbtob / 2)
              
              By(5) = valYbtob + valYbtob / 2
              By(6) = valYbtob / 2
              By(7) = -(valYbtob / 2)
              By(8) = -(valYbtob + valYbtob / 2)
          
      End Select
Else
      Select Case valType
          Case "I"
              BX(1) = 0: BX(2) = 0
              By(1) = valYbtob / 2: By(2) = -(valYbtob / 2)
          Case "II"
              BX(1) = valXbtob / 2: BX(2) = valXbtob / 2
              BX(3) = -(valXbtob / 2): BX(4) = -(valXbtob / 2)
              
              By(1) = valYbtob / 2: By(2) = -(valYbtob / 2)
              By(3) = valYbtob / 2: By(4) = -(valYbtob / 2)
          Case "III"
              BX(1) = valXbtob: BX(2) = 0: BX(3) = -1 * valXbtob
              BX(4) = valXbtob: BX(5) = 0: BX(6) = -1 * valXbtob
              
              By(1) = valYbtob / 2: By(2) = valYbtob / 2: By(3) = valYbtob / 2
              By(4) = -(valYbtob / 2): By(5) = -(valYbtob / 2): By(6) = -(valYbtob / 2)
          Case "IV"
              BX(1) = valXbtob + valXbtob / 2
              BX(2) = valXbtob / 2
              BX(3) = -(valXbtob / 2)
              BX(4) = -(valXbtob + valXbtob / 2)
              
              BX(5) = valXbtob + valXbtob / 2
              BX(6) = valXbtob / 2
              BX(7) = -(valXbtob / 2)
              BX(8) = -(valXbtob + valXbtob / 2)
              
              
              By(1) = valYbtob / 2: By(2) = valYbtob / 2: By(3) = valYbtob / 2: By(4) = valYbtob / 2
              By(5) = -(valYbtob / 2): By(6) = -(valYbtob / 2): By(7) = -(valYbtob / 2): By(8) = -(valYbtob / 2)
              

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
            Print #1, "vert" & j & " = " & Format(Nx(i, j), "0.000") & ", " & Format(Ny(i, j), "0.000") & _
                       "," & Format(valPt + valNutHei, "0.000") & ","
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

Public Function convert(valOrigin As Single, valinput As String, valoutput As String) As Single

'valinput : Unit System에서 선택한 Unit
'valoutput : Output에서 사용되는 Unit
Select Case valinput
    Case "mm"
        Select Case valoutput
            Case "mm"
                convert = valOrigin
            Case "cm"
                convert = valOrigin / 10
            Case "m"
                convert = valOrigin / 1000
            Case "inch"
                convert = valOrigin / 25.4
            Case "feet"
                convert = valOrigin / (25.4 * 12)
        End Select
'    Case "cm"
'        Select Case valoutput
'            Case "mm"
'                convert = valOrigin * 10
'            Case "cm"
'                convert = valOrigin
'            Case "m"
'                convert = valOrigin / 100
'            Case "inch"
'                convert = valOrigin / 2.54
'            Case "feet"
'                convert = valOrigin / (2.54 * 12)
'        End Select
    Case "m"
        Select Case valoutput
            Case "mm"
                convert = valOrigin * 1000
            Case "cm"
                convert = valOrigin * 100
            Case "m"
                convert = valOrigin
            Case "inch"
                convert = valOrigin / 0.0254
            Case "feet"
                convert = valOrigin / (0.0254 * 12)
        End Select
    Case "inch"
        Select Case valoutput
            Case "mm"
                convert = valOrigin * 25.4
            Case "cm"
                convert = valOrigin * 2.54
            Case "m"
                convert = valOrigin * 0.0254
            Case "inch"
                convert = valOrigin
            Case "feet"
                convert = valOrigin / 12
        End Select
    Case "feet"
        Select Case valoutput
            Case "mm"
                convert = valOrigin * 25.4 * 12
            Case "cm"
                convert = valOrigin * 2.54 * 12
            Case "m"
                convert = valOrigin * 0.0254 * 12
            Case "inch"
                convert = valOrigin * 12
            Case "feet"
                convert = valOrigin
        End Select
        
End Select

End Function



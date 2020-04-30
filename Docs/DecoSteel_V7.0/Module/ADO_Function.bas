Attribute VB_Name = "ADO_Function"
Public adoConnection As ADODB.Connection
Public adoRecordset As ADODB.Recordset
Public adoConnection1 As ADODB.Connection
Public adoRecordset1 As ADODB.Recordset

Public Sub ADO_Connection()

Dim ConnectionString As String
Dim pathName As String

Set adoConnection = New ADODB.Connection
Set adoRecordset = New ADODB.Recordset
Set adoConnection1 = New ADODB.Connection
Set adoRecordset1 = New ADODB.Recordset

pathName = App.Path & "\library\steelmember.mdb"

ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                    "Data Source = " & pathName

adoConnection.Open ConnectionString

pathName = App.Path & "\library\SCS.mdb"

ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                    "Data Source = " & pathName

adoConnection1.Open ConnectionString
End Sub

Public Sub ADO_disConnection()
adoConnection.Close
adoConnection1.Close
Set adoConnection = Nothing
Set adoConnection1 = Nothing

End Sub
Public Sub Query_AddComb_function(combList As ComboBox, CodeName As String, Flag As String)

Dim Sql As String
combList.Clear

Sql = "Select * from code_" & CodeName & _
      " where member_type = '" & Flag & "' order by member_NO"
Set adoRecordset = adoConnection.Execute(Sql)

Do Until adoRecordset.EOF
    combList.AddItem adoRecordset!Member_Name
    adoRecordset.MoveNext
Loop

adoRecordset.Close

Set adoRecordset = Nothing

End Sub

Public Sub Query_AddList2_function(Flag As Integer, ListBox As Control, valSQL As String)

ListBox.Clear

If Flag = 0 Then
    Set adoRecordset = adoConnection.Execute(valSQL)
    Do Until adoRecordset.EOF
        ListBox.AddItem adoRecordset(0)
        adoRecordset.MoveNext
    Loop

    adoRecordset.Close
    
    Set adoRecordset = Nothing

Else
    Set adoRecordset1 = adoConnection1.Execute(valSQL)
    Do Until adoRecordset1.EOF
        ListBox.AddItem adoRecordset1(0)
        adoRecordset1.MoveNext
    Loop

    adoRecordset1.Close
    
    Set adoRecordset1 = Nothing
End If


End Sub

Public Sub Query_AddList_function(Flag As Integer, ListBox As Control, valSQL As String)

ListBox.Clear

If Flag = 0 Then
    Set adoRecordset = adoConnection.Execute(valSQL)
    
    Do Until adoRecordset.EOF
        ListBox.AddItem adoRecordset!Member_Name
        adoRecordset.MoveNext
    Loop
    
    adoRecordset.Close
    
    Set adoRecordset = Nothing
Else
    Set adoRecordset1 = adoConnection1.Execute(valSQL)
    
    Do Until adoRecordset1.EOF
        ListBox.AddItem adoRecordset1!Member_Name
        adoRecordset1.MoveNext
    Loop
    
    adoRecordset1.Close
    
    Set adoRecordset1 = Nothing
End If
End Sub

Public Sub Query_MemberData(CodeName As String, Mname As String, _
                            ByRef Height As Double, ByRef Width As Double, ByRef Tf As Double, _
                            ByRef Tw As Double)
Dim Sql As String
Dim TempFlag As String


Sql = "select * from code_" & CodeName & _
      " where member_name = '" & Mname & "'"

Set adoRecordset = adoConnection.Execute(Sql)

With adoRecordset
    If !D = "" Then
        Height = 0
    Else
        Height = CDbl(!D)
    End If
    
    If !Bf = "" Then
        Width = 0
    Else
        Width = CDbl(!Bf)
    End If
    
    If !Tf = "" Then
        Tf = 0
    Else
        Tf = CDbl(!Tf)
    End If
    
    If !Tw = "" Then
        Tw = 0
    Else
        Tw = CDbl(!Tw)
    End If
    
End With


adoRecordset.Close
Set adoRecordset = Nothing

End Sub
Public Sub gsubSSADOQuery(Flag As Integer, Query As String, ss As Control, Optional Col, Optional Row)
    
    Dim retData As ADODB.Recordset

    Dim i As Long
    Dim OldMousePointer As Integer
    Dim lngRow As Long
    Dim lngCol As Long

On Error GoTo SelectError

    OldMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    DoEvents

    If (IsMissing(Col)) Then
        lngRow = 0
        lngCol = 1
        Call gsubSS_SetMax(ss, 0, ss.MaxCols)
    Else
        lngRow = Row - 1
        lngCol = Col ' - 1
        Call gsubSS_Clear(ss, , Row, Col, ss.MaxRows, ss.MaxCols)
        ss.MaxRows = lngRow
    End If
    
    If Flag = 0 Then
        Set retData = adoConnection.Execute(Query)
    Else
        Set retData = adoConnection1.Execute(Query)
    End If
    
    ss.ReDraw = False

    Do Until retData.EOF
        lngRow = lngRow + 1
        If lngRow > ss.MaxRows Then ss.MaxRows = lngRow
        ss.Row = lngRow
        ss.Col = 0
        ss.Text = ""

        For i = 0 To retData.Fields.Count - 1 'ss.MaxCols - 1          'Query 결과 Browse
            
            ss.Col = lngCol + i
            
            If IsNull(retData.Fields(i)) Then
                ss.Text = ""
            Else
                If retData(i).Type = 10 Or retData(i).Type = 12 Then
                    ss.Text = retData(i)
                Else
                    ss.Text = LTrim(CStr(retData(i)))
                End If
            End If

        Next i

        retData.MoveNext

        If ss.MaxRows >= ss.VisibleRows Then DoEvents
        If ss.MaxRows = ss.VisibleRows Then ss.ReDraw = True
    Loop
    If ss.MaxRows < ss.VisibleRows Then ss.ReDraw = True
    retData.Close
    Set retData = Nothing
    Screen.MousePointer = OldMousePointer
    Exit Sub

SelectError:
    Call gsubSS_SetMax(ss, ss.RowsFrozen, ss.MaxCols)     'Sheet Clear
    Set retData = Nothing
    Screen.MousePointer = OldMousePointer
    MsgBox "자료검색중 오류가 발생했습니다." & Chr(13) & Chr(13) & Chr(10) & _
            Query & Chr(13) & Chr(13) & Chr(10) & _
           Err & ":" & Error, vbCritical, "ERROR"
End Sub

Public Function DataCheck(valName As String, valSQL As String) As Boolean
Dim xSQL As String
Dim retData As ADODB.Recordset

    
    Set retData = adoConnection1.Execute(valSQL)
    
    Do Until retData.EOF
        If valName = CStr(Trim(retData!Member_Name)) Then
            DataCheck = True
        Else
            DataCheck = False
        End If
        retData.MoveNext
    Loop
    
    Set retData = Nothing
End Function
Public Function DataCheck_MC(valName As String, valType As String, valSQL As String) As Boolean
Dim xSQL As String
Dim retData As ADODB.Recordset

    
    Set retData = adoConnection1.Execute(valSQL)
    
    Do Until retData.EOF
        If valName = CStr(Trim(retData!Member_Name)) And valType = CStr(Trim(retData!Type)) Then
            DataCheck_MC = True
        Else
            DataCheck_MC = False
        End If
        retData.MoveNext
    Loop
    
    Set retData = Nothing
End Function


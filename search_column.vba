Sub Btn_doSearch_Click()

'-----------------------------------------'
sheet_name = "collection"
column_ids = 1 ' first column is reserved for id-numbers
col2search = 2 ' 2nd column 2 be searched
max_result = 15
'-----------------------------------------'

'Input
input_term = InputBox("Please enter name or term to search for.", "Search")

If input_term = "" Then
    MsgBox "Nothing entered. Please try again and enter a name or term to search for.", vbOKOnly + vbInformation, "Attention"
    Exit Sub
End If
'Input

If Worksheets(sheet_name).FilterMode Then Worksheets(sheet_name).ShowAllData ' reset Filter

Dim SBegriff
SBegriff = input_term

    Dim c
    Dim c1st
    Dim ic
    ic = 1
    Dim ice
    ice = 1
    Dim i
    i = 1
    
    Dim ErgebWert1()
    Dim ErgebWert2()
    Dim AusgabeMSG

    '***********************************
    With Worksheets(sheet_name).Columns(col2search)
        Set c = .Find(SBegriff, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) 'LookIn:=xlValues ignores Formula in cells
        If Not c Is Nothing Then
            c1st = c.Address
            Do
                ice = ice + 1
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> c1st
            ReDim ErgebWert1(ice)
            ReDim ErgebWert2(ice)
            Do
                ErgebWert1(ic) = Worksheets(sheet_name).Cells(c.Row, col2search).Value 'read name
                ErgebWert2(ic) = Worksheets(sheet_name).Cells(c.Row, column_ids).Value 'read ID
            ic = ic + 1
            Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> c1st
            Do
                AusgabeMSG = AusgabeMSG & i & ". " & ErgebWert1(i) & " (ID: " & ErgebWert2(i) & ")" & vbNewLine
                i = i + 1
            Loop While i <> ic
        End If
    End With
    '***********************************

If i = 1 Then
            MsgBox "No matching term found." & vbNewLine & vbNewLine & "Please varify manually.", vbOKOnly + vbInformation, "Attention"
ElseIf i > 1 And i < max_result + 1 Then
            MsgBox "Found one or more similar entries:" & vbNewLine & vbNewLine & AusgabeMSG, vbOKOnly + vbInformation, "Attention"
ElseIf i > max_result Then
            MsgBox "More than ''" & max_result & "'' similar entries with ''" & SBegriff & "'' found in working sheet." & vbNewLine & vbNewLine & "Please check manually.", vbOKOnly + vbInformation, "Attention"
End If

End Sub

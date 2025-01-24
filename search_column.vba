Private Sub Cmd_Btn_CheckName_Click()

If Worksheets("collection").FilterMode Then Worksheets("collection").ShowAllData ' reset Filter


'Bitte einen Suchbegriff eingeben
If UnternehmenEingabe.UN_Name.Value = "" Then
    MsgBox "Please enter a name or term to search for.", vbOKOnly + vbInformation, "Attention"
    UnternehmenEingabe.UN_Name.SetFocus
    Exit Sub
End If
'Bitte einen Suchbegriff eingeben


Dim SBegriff
SBegriff = UnternehmenEingabe.UN_Name.Value

    Dim c
    Dim cErst
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
    With Worksheets("collection").Columns(4)
    
        Set c = .Find(SBegriff, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False) 'LookIn:=xlValues ignores Formula in cells

        If Not c Is Nothing Then
        
            cErst = c.Address
            
            Do
                ice = ice + 1
                Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> cErst
            
            ReDim ErgebWert1(ice)
            ReDim ErgebWert2(ice)
            
            Do
                ErgebWert1(ic) = Worksheets("collection").Cells(c.Row, 4).Value 'read name
                ErgebWert2(ic) = Worksheets("collection").Cells(c.Row, 1).Value 'read ID

            ic = ic + 1
            Set c = .FindNext(c)
            Loop While Not c Is Nothing And c.Address <> cErst
            
            Do
                AusgabeMSG = AusgabeMSG & i & ". " & ErgebWert1(i) & " (ID: " & ErgebWert2(i) & ")" & vbNewLine
                i = i + 1
            Loop While i <> ic
        
        End If
    
    End With
    '***********************************

'If ic = 1 Then
'    MsgBox "Not found.", vbOKOnly + vbInformation, "Attention"
'    Else: UnternehmenEingabe.Ergebnis_Auswahl.Text = "Show results / choose entry"
'End If

If i = 1 Then
            MsgBox "No matching Name found." & vbNewLine & vbNewLine & "Please varify manually.", vbOKOnly + vbInformation, "Attention"
ElseIf i > 1 And i < 16 Then
            MsgBox "Found one or more similar entries:" & vbNewLine & vbNewLine & AusgabeMSG, vbOKOnly + vbInformation, "Attention"
ElseIf i > 15 Then
            MsgBox "More than 15 similar entries with ''" & SBegriff & "'' found in working sheet." & vbNewLine & vbNewLine & "Please check to avoid redundancies.", vbOKOnly + vbInformation, "Attention"
End If

End Sub
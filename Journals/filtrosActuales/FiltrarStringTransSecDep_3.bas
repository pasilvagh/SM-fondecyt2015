Attribute VB_Name = "Módulo1"
Sub FiltrarStringBusqueda()
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim str5 As String
Dim str6 As String
Dim str7 As String
Dim str8 As String
Dim str9 As String
Dim str10 As String
Dim str11 As String
Dim str12 As String
Dim str13 As String
Dim str14 As String
Dim str15 As String
Dim str16 As String
Dim str17 As String
Dim str18 As String
Dim str19 As String
Dim str20 As String
Dim str21 As String
Dim str22 As String
Dim str23 As String
Dim str24 As String
Dim str25 As String
Dim str26 As String
Dim str27 As String
Dim str28 As String
Dim str29 As String
Dim str30 As String
Dim str31 As String
Dim str32 As String
Dim str33 As String
Dim str34 As String

Dim cond1 As Boolean
Dim cond2 As Boolean
Dim cond3 As Boolean
Dim cond4 As Boolean
Dim cond5 As Boolean


str1 = UCase("Software") 
str2 = UCase("Design")
str3 = UCase("Engineer")
str4 = UCase("Develop")
str5 = UCase("Securit")
str6 = UCase("Privacy")
str7 = UCase("Integrity")
str8 = UCase("Confidential")
str9 = UCase("Availab")
str10 = UCase("Accountab")
str11 = UCase("Threat")
str12 = UCase("Risk")
str13 = UCase("Attack")
str14 = UCase("Requirement")
str15 = UCase("Vulnerabil")
str16 = UCase("Indentif")
str17 = UCase("Mitig")
str18 = UCase("Minimiz")
str19 = UCase("Elicit")
str20 = UCase("Enumer")
str21 = UCase("Review")
str22 = UCase("Assur")
str23 = UCase("Model")
str24 = UCase("Metric")
str25 = UCase("Guideline")
str26 = UCase("Checklist")
str27 = UCase("Template")
str28 = UCase("Approach")
str29 = UCase("Strateg")
str30 = UCase("Method")
str31 = UCase("Methodolog")
str32 = UCase("Tool")
str33 = UCase("Technique")
str34 = UCase("Heuristic")



Dim i As Integer

For i = 3 To 483
    texto = UCase(Cells(i, 7).Value)

    ''("Software" OR "Design" OR "Engineer" OR "Develop") AND
    If (InStr(texto,str1) OR InStr(texto,str2) OR InStr(texto,str3) OR InStr(texto,str4)) Then
        cond1 = True
    Else
        cond1 = False
    End If

    ''("Securit" OR "Privacy" OR "Integrity" OR "Confidential" OR "Availabil" OR "Accountabil") 
    If (InStr(texto,str5) OR InStr(texto,str6) OR InStr(texto,str7) OR InStr(texto,str8) OR InStr(texto,str9) OR InStr(texto,str10)) Then
        cond2 = True
    Else
        cond2 = False
    End If

    ''("Threat" OR "Risk" OR "Attack" OR "Requirement" OR "Vulnerabil")
    If (InStr(texto, str11) OR InStr(texto, str12) OR InStr(texto, str13) OR InStr(texto, str14) OR InStr(texto, str15)) Then
        cond3 = True
    Else
        cond3 = False
    End If

    ''("Identif" OR "Mitig" OR "Minimiz" OR "Elicit" OR "Enumer" OR "Review" OR "Assur")
    If (InStr(texto, str16) OR InStr(texto, str17) OR InStr(texto, str18) OR InStr(texto, str19) OR InStr(texto, str20) OR InStr(texto, str21) OR InStr(texto, str22)) Then
        cond4 = True
    Else
        cond4 = False
    End If

    ''("model" OR "metric" OR "guideline" OR "checklist" OR "template" OR "approach" OR "strateg" OR "method" OR "methodolog" OR "tool" OR "technique" OR "heuristic")
    If (InStr(texto, str23) OR InStr(texto, str24) OR InStr(texto, str25) OR InStr(texto, str26) OR InStr(texto, str27) OR InStr(texto, str28) OR InStr(texto, str29) OR InStr(texto, str30) OR InStr(texto, str31) OR InStr(texto, str32) OR InStr(texto, str33) OR InStr(texto, str34)) Then
        cond5 = True
    Else
        cond5 = False
    End If

    If (cond1 And (cond2 And cond3 And cond4) And cond5) Then
        Cells(i, 11).Value = True
    Else
        Cells(i, 11).Value = False
    End If
                          
Next i
End Sub

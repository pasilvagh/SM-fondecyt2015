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
str3 = UCase("Engineering")
str4 = UCase("Development")
str5 = UCase("Security")
str6 = UCase("Privacy")
str7 = UCase("Integrity")
str8 = UCase("Confidentiallity")
str9 = UCase("Availability")
str10 = UCase("Accountability")
str11 = UCase("Threat")
str12 = UCase("Risk")
str13 = UCase("Attack")
str14 = UCase("Requirement")
str15 = UCase("Vulnerability")
str16 = UCase("Indentification")
str17 = UCase("Mitigation")
str18 = UCase("Minimize")
str19 = UCase("Elicitation")
str20 = UCase("Enumeration")
str21 = UCase("Review")
str22 = UCase("Assurance")
str23 = UCase("Model")
str24 = UCase("Metric")
str25 = UCase("Guideline")
str26 = UCase("Checklist")
str27 = UCase("Template")
str28 = UCase("Approach")
str29 = UCase("Strategy")
str30 = UCase("Method")
str31 = UCase("Methodology")
str32 = UCase("Tool")
str33 = UCase("Technique")
str34 = UCase("Heuristic")



Dim i As Integer

For i = 3 To 77
    texto = UCase(Cells(i, 7).Value)

    ''(("Software") AND ("Design" OR "Engineering" OR "Development")) AND       
    If (InStr(texto,str1) AND (InStr(texto,str2) OR InStr(texto,str3) OR InStr(texto,str4))) Then
        cond1 = True
    Else
        cond1 = False
    End If

    ''("Security" OR "Privacy" OR "Integrity" OR "Confidentiality" OR "Availability" OR "Accountability")
    If (InStr(texto,str5) OR InStr(texto,str6) OR InStr(texto,str7) OR InStr(texto,str8) OR InStr(texto,str9) OR InStr(texto,str10)) Then
        cond2 = True
    Else
        cond2 = False
    End If

    ''("Threat" OR "Risk" OR "Attack" OR "Requirement" OR "Vulnerability")
    If (InStr(texto, str11) OR InStr(texto, str12) OR InStr(texto, str13) OR InStr(texto, str14) OR InStr(texto, str15)) Then
        cond3 = True
    Else
        cond3 = False
    End If

    ''("Identification" OR "Mitigation" OR "Minimize" OR "Elicitation" OR "Enumeration" OR "Review" OR "Assurance")
    If (InStr(texto, str16) OR InStr(texto, str17) OR InStr(texto, str18) OR InStr(texto, str19) OR InStr(texto, str20) OR InStr(texto, str21) OR InStr(texto, str22)) Then
        cond4 = True
    Else
        cond4 = False
    End If

    ''("model" OR "metric" OR "guideline" OR "checklist" OR "template" OR "approach" OR "strategy" OR "method" OR "methodology" OR "tool" OR "technique" OR "heuristic")
    If (InStr(texto, str23) OR InStr(texto, str24) OR InStr(texto, str25) OR InStr(texto, str26) OR InStr(texto, str27) OR InStr(texto, str28) OR InStr(texto, str29) OR InStr(texto, str30) OR InStr(texto, str31) OR InStr(texto, str32) OR InStr(texto, str33) OR InStr(texto, str34)) Then
        cond5 = True
    Else
        cond5 = False
    End If

    If (cond1 And (cond2 And cond3 And cond4) And cond5) Then
        Cells(i, 10).Value = True
    Else
        Cells(i, 10).Value = False
    End If
                          
Next i
End Sub

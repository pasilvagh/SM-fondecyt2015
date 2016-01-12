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
Dim cond1 As Boolean
Dim cond2 As Boolean
Dim cond3 As Boolean
Dim cond4 As Boolean
Dim cond5 As Boolean

str1 = UCase("Software")
str1_1 = UCase("Service")
str2 = UCase("Design")
str3 = UCase("Engineering")
str3_1 = UCase("Develop")
str4 = UCase("System")
str5 = UCase("Threat")
str5_1 = UCase("Risk")
str6 = UCase("Attack")
str7 = UCase("Requirement")
str8 = UCase("Vulnerab")
str9 = UCase("Ident")
str10 = UCase("Mitigat")
str10_1 = UCase("Minimize")
str11 = UCase("Elicit")
str12 = UCase("Enum")
str13 = UCase("Review")
str13_1 = UCase("Assur")
str14 = UCase("Secur")
str15 = UCase("Priva")
str16 = UCase("Integrit")
str17 = UCase("Confident")
str18 = UCase("Availab")
str19 = UCase("Account")

Dim i As Integer

For i = 3 To 1714
    texto = UCase(Cells(i, 7).Value)

        ''("Software" or "Service" "System" or )
    If ((InStr(texto, str1) Or InStr(texto, str1_1) Or InStr(texto, str4))) Then
        cond1 = True
    Else
        cond1 = False
    End If

        ''("Design" or "Engineering" or "Develop")
    If (InStr(texto, str2) Or InStr(texto, str3) Or InStr(texto, str3_1)) Then
        cond2 = True
    Else
        cond2 = False
    End If

    ''("Threat" or "Risk" or "Attack" or "Requirement" or "Vulnerab")
    If (InStr(texto, str5) Or InStr(texto, str5_1) Or InStr(texto, str6) Or InStr(texto, str7) Or InStr(texto, str8)) Then
        cond3 = True
    Else
        cond3 = False
    End If

    ''("Ident" or "Mitigat" or "Minimize" or "Elicit" or "Enum" or "Review" or "Assur")
    If (InStr(texto, str9) Or InStr(texto, str10) Or InStr(texto, str10_1) Or InStr(texto, str11) Or InStr(texto, str12) Or InStr(texto, str13) Or InStr(texto, str13_1)) Then
        cond4 = True
    Else
        cond4 = False
    End If

    ''("Secur" or "Priva" or "Integrit" or "Confident" or "Availab" or "Account")
    If (InStr(texto, str14) Or InStr(texto, str15) Or InStr(texto, str16) Or InStr(texto, str17) Or InStr(texto, str18) Or InStr(texto, str19)) Then
        cond5 = True
    Else
        cond5 = False
    End If

    If (cond1 And cond2 And cond3 And cond4 And cond5) Then
        Cells(i, 9).Value = True
    Else
        Cells(i, 9).Value = False
    End If
                          
Next i
End Sub

Attribute VB_Name = "ModuloValidacoes"
' ==============================================================================
' MÓDULO DE VALIDAÇÕES E FORMATAÇÕES
' Organizador de Declaração de Imposto de Renda
' Desenvolvido por: Marcus Vasconcellos - DIO Bootcamp Santander 2025
' ==============================================================================

Option Explicit

' ==============================================================================
' FUNÇÃO: ValidarCPF
' DESCRIÇÃO: Valida CPF verificando os dígitos verificadores
' PARÂMETROS: cpf (String) - CPF a ser validado (somente números)
' RETORNO: Boolean - True se válido, False se inválido
' ==============================================================================
Public Function ValidarCPF(cpf As String) As Boolean
    Dim i As Integer
    Dim soma As Integer
    Dim resto As Integer
    Dim digito1 As Integer
    Dim digito2 As Integer
    Dim cpfNumeros As String
    
    ' Remove caracteres não numéricos
    cpfNumeros = RemoverNaoNumericos(cpf)
    
    ' Verifica se tem 11 dígitos
    If Len(cpfNumeros) <> 11 Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Verifica se todos os dígitos são iguais (CPF inválido)
    If cpfNumeros = String(11, Mid(cpfNumeros, 1, 1)) Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Cálculo do primeiro dígito verificador
    soma = 0
    For i = 1 To 9
        soma = soma + CInt(Mid(cpfNumeros, i, 1)) * (11 - i)
    Next i
    
    resto = (soma * 10) Mod 11
    If resto = 10 Or resto = 11 Then resto = 0
    digito1 = resto
    
    ' Verifica primeiro dígito
    If digito1 <> CInt(Mid(cpfNumeros, 10, 1)) Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Cálculo do segundo dígito verificador
    soma = 0
    For i = 1 To 10
        soma = soma + CInt(Mid(cpfNumeros, i, 1)) * (12 - i)
    Next i
    
    resto = (soma * 10) Mod 11
    If resto = 10 Or resto = 11 Then resto = 0
    digito2 = resto
    
    ' Verifica segundo dígito
    If digito2 <> CInt(Mid(cpfNumeros, 11, 1)) Then
        ValidarCPF = False
        Exit Function
    End If
    
    ValidarCPF = True
End Function

' ==============================================================================
' FUNÇÃO: ValidarCNPJ
' DESCRIÇÃO: Valida CNPJ verificando os dígitos verificadores
' PARÂMETROS: cnpj (String) - CNPJ a ser validado
' RETORNO: Boolean - True se válido, False se inválido
' ==============================================================================
Public Function ValidarCNPJ(cnpj As String) As Boolean
    Dim i As Integer
    Dim soma As Integer
    Dim resto As Integer
    Dim digito1 As Integer
    Dim digito2 As Integer
    Dim cnpjNumeros As String
    Dim multiplicadores1 As Variant
    Dim multiplicadores2 As Variant
    
    multiplicadores1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    multiplicadores2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    ' Remove caracteres não numéricos
    cnpjNumeros = RemoverNaoNumericos(cnpj)
    
    ' Verifica se tem 14 dígitos
    If Len(cnpjNumeros) <> 14 Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Cálculo do primeiro dígito
    soma = 0
    For i = 0 To 11
        soma = soma + CInt(Mid(cnpjNumeros, i + 1, 1)) * multiplicadores1(i)
    Next i
    
    resto = soma Mod 11
    If resto < 2 Then
        digito1 = 0
    Else
        digito1 = 11 - resto
    End If
    
    If digito1 <> CInt(Mid(cnpjNumeros, 13, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Cálculo do segundo dígito
    soma = 0
    For i = 0 To 12
        soma = soma + CInt(Mid(cnpjNumeros, i + 1, 1)) * multiplicadores2(i)
    Next i
    
    resto = soma Mod 11
    If resto < 2 Then
        digito2 = 0
    Else
        digito2 = 11 - resto
    End If
    
    If digito2 <> CInt(Mid(cnpjNumeros, 14, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ValidarCNPJ = True
End Function

' ==============================================================================
' FUNÇÃO: FormatarCPF
' DESCRIÇÃO: Formata CPF no padrão ###.###.###-##
' PARÂMETROS: cpf (String) - CPF sem formatação
' RETORNO: String - CPF formatado
' ==============================================================================
Public Function FormatarCPF(cpf As String) As String
    Dim cpfNumeros As String
    
    cpfNumeros = RemoverNaoNumericos(cpf)
    
    If Len(cpfNumeros) = 11 Then
        FormatarCPF = Mid(cpfNumeros, 1, 3) & "." & _
                     Mid(cpfNumeros, 4, 3) & "." & _
                     Mid(cpfNumeros, 7, 3) & "-" & _
                     Mid(cpfNumeros, 10, 2)
    Else
        FormatarCPF = cpf
    End If
End Function

' ==============================================================================
' FUNÇÃO: FormatarCNPJ
' DESCRIÇÃO: Formata CNPJ no padrão ##.###.###/####-##
' ==============================================================================
Public Function FormatarCNPJ(cnpj As String) As String
    Dim cnpjNumeros As String
    
    cnpjNumeros = RemoverNaoNumericos(cnpj)
    
    If Len(cnpjNumeros) = 14 Then
        FormatarCNPJ = Mid(cnpjNumeros, 1, 2) & "." & _
                      Mid(cnpjNumeros, 3, 3) & "." & _
                      Mid(cnpjNumeros, 6, 3) & "/" & _
                      Mid(cnpjNumeros, 9, 4) & "-" & _
                      Mid(cnpjNumeros, 13, 2)
    Else
        FormatarCNPJ = cnpj
    End If
End Function

' ==============================================================================
' FUNÇÃO: FormatarTelefone
' DESCRIÇÃO: Formata telefone (##) ####-#### ou (##) #####-####
' ==============================================================================
Public Function FormatarTelefone(telefone As String) As String
    Dim telNumeros As String
    
    telNumeros = RemoverNaoNumericos(telefone)
    
    If Len(telNumeros) = 10 Then
        ' Telefone fixo
        FormatarTelefone = "(" & Mid(telNumeros, 1, 2) & ") " & _
                          Mid(telNumeros, 3, 4) & "-" & _
                          Mid(telNumeros, 7, 4)
    ElseIf Len(telNumeros) = 11 Then
        ' Celular
        FormatarTelefone = "(" & Mid(telNumeros, 1, 2) & ") " & _
                          Mid(telNumeros, 3, 5) & "-" & _
                          Mid(telNumeros, 8, 4)
    Else
        FormatarTelefone = telefone
    End If
End Function

' ==============================================================================
' FUNÇÃO: FormatarCEP
' DESCRIÇÃO: Formata CEP no padrão #####-###
' ==============================================================================
Public Function FormatarCEP(cep As String) As String
    Dim cepNumeros As String
    
    cepNumeros = RemoverNaoNumericos(cep)
    
    If Len(cepNumeros) = 8 Then
        FormatarCEP = Mid(cepNumeros, 1, 5) & "-" & Mid(cepNumeros, 6, 3)
    Else
        FormatarCEP = cep
    End If
End Function

' ==============================================================================
' FUNÇÃO: RemoverNaoNumericos
' DESCRIÇÃO: Remove todos os caracteres não numéricos de uma string
' ==============================================================================
Private Function RemoverNaoNumericos(texto As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim caractere As String
    
    resultado = ""
    For i = 1 To Len(texto)
        caractere = Mid(texto, i, 1)
        If IsNumeric(caractere) Then
            resultado = resultado & caractere
        End If
    Next i
    
    RemoverNaoNumericos = resultado
End Function

' ==============================================================================
' SUBROTINA: AplicarValidacaoAba
' DESCRIÇÃO: Aplica validações automáticas na aba TITULAR
' ==============================================================================
Public Sub AplicarValidacaoAba()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("TITULAR")
    
    ' Exemplo de uso - adaptar conforme células da planilha
    ' Assumindo que CPF está na célula D4
    
    Application.EnableEvents = False
    
    With ws
        ' Aplicar formatação de CPF
        If Not IsEmpty(.Range("D4")) Then
            .Range("D4").Value = FormatarCPF(.Range("D4").Value)
        End If
        
        ' Aplicar formatação de telefone
        If Not IsEmpty(.Range("D10")) Then
            .Range("D10").Value = FormatarTelefone(.Range("D10").Value)
        End If
        
        ' Aplicar formatação de celular
        If Not IsEmpty(.Range("D11")) Then
            .Range("D11").Value = FormatarTelefone(.Range("D11").Value)
        End If
        
        ' Aplicar formatação de CEP
        If Not IsEmpty(.Range("D9")) Then
            .Range("D9").Value = FormatarCEP(.Range("D9").Value)
        End If
    End With
    
    Application.EnableEvents = True
    
    MsgBox "Formatações aplicadas com sucesso!", vbInformation, "LION APP"
End Sub

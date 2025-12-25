Attribute VB_Name = "ModuloCalculos"
' ==============================================================================
' MÓDULO DE CÁLCULOS DE IMPOSTO DE RENDA
' Organizador de Declaração de Imposto de Renda
' Desenvolvido por: Marcus Vasconcellos - DIO Bootcamp Santander 2025
' ==============================================================================

Option Explicit

' Constantes da Tabela Progressiva 2025
Private Const FAIXA1_LIMITE As Double = 2259.2
Private Const FAIXA2_LIMITE As Double = 2826.65
Private Const FAIXA3_LIMITE As Double = 3751.05
Private Const FAIXA4_LIMITE As Double = 4664.68

Private Const ALIQ1 As Double = 0
Private Const ALIQ2 As Double = 0.075
Private Const ALIQ3 As Double = 0.15
Private Const ALIQ4 As Double = 0.225
Private Const ALIQ5 As Double = 0.275

Private Const DEDUZIR1 As Double = 0
Private Const DEDUZIR2 As Double = 169.44
Private Const DEDUZIR3 As Double = 381.44
Private Const DEDUZIR4 As Double = 662.77
Private Const DEDUZIR5 As Double = 896#

' Outras constantes
Private Const DEDUCAO_DEPENDENTE As Double = 2275.08
Private Const DEDUCAO_EDUCACAO_LIMITE As Double = 3561.5
Private Const DESCONTO_SIMPLIFICADO_PERC As Double = 0.2
Private Const DESCONTO_SIMPLIFICADO_LIMITE As Double = 16754.34

' ==============================================================================
' FUNÇÃO: CalcularImpostoMensal
' DESCRIÇÃO: Calcula imposto mensal com base na tabela progressiva
' PARÂMETROS: baseCalculo (Double) - Base de cálculo mensal
' RETORNO: Double - Imposto calculado
' ==============================================================================
Public Function CalcularImpostoMensal(baseCalculo As Double) As Double
    Dim imposto As Double
    
    If baseCalculo <= FAIXA1_LIMITE Then
        imposto = 0
    ElseIf baseCalculo <= FAIXA2_LIMITE Then
        imposto = (baseCalculo * ALIQ2) - DEDUZIR2
    ElseIf baseCalculo <= FAIXA3_LIMITE Then
        imposto = (baseCalculo * ALIQ3) - DEDUZIR3
    ElseIf baseCalculo <= FAIXA4_LIMITE Then
        imposto = (baseCalculo * ALIQ4) - DEDUZIR4
    Else
        imposto = (baseCalculo * ALIQ5) - DEDUZIR5
    End If
    
    If imposto < 0 Then imposto = 0
    CalcularImpostoMensal = imposto
End Function

' ==============================================================================
' FUNÇÃO: CalcularBaseCalculo
' DESCRIÇÃO: Calcula base de cálculo mensal
' ==============================================================================
Public Function CalcularBaseCalculo(rendimentoBruto As Double, _
                                   numeroDependentes As Integer, _
                                   deducaoINSS As Double, _
                                   Optional outrasDeduc As Double = 0) As Double
    
    Dim deducaoDependentes As Double
    Dim baseCalculo As Double
    
    deducaoDependentes = numeroDependentes * DEDUCAO_DEPENDENTE
    baseCalculo = rendimentoBruto - deducaoINSS - deducaoDependentes - outrasDeduc
    
    If baseCalculo < 0 Then baseCalculo = 0
    CalcularBaseCalculo = baseCalculo
End Function

' ==============================================================================
' FUNÇÃO: CalcularAliquotaEfetiva
' DESCRIÇÃO: Calcula alíquota efetiva do imposto
' ==============================================================================
Public Function CalcularAliquotaEfetiva(impostoDevido As Double, _
                                        rendimentoBruto As Double) As Double
    If rendimentoBruto > 0 Then
        CalcularAliquotaEfetiva = (impostoDevido / rendimentoBruto) * 100
    Else
        CalcularAliquotaEfetiva = 0
    End If
End Function

' ==============================================================================
' FUNÇÃO: CompararDeducaoCompleta
Simplificada
' DESCRIÇÃO: Compara e recomenda modelo de declaração
' RETORNO: String - "COMPLETA" ou "SIMPLIFICADA"
' ==============================================================================
Public Function CompararDeducaoCompletaSimplificada(rendimentoBruto As Double, _
                                                   deducoesCompleta As Double) As String
    Dim descontoSimplificado As Double
    
    descontoSimplificado = rendimentoBruto * DESCONTO_SIMPLIFICADO_PERC
    If descontoSimplificado > DESCONTO_SIMPLIFICADO_LIMITE Then
        descontoSimplificado = DESCONTO_SIMPLIFICADO_LIMITE
    End If
    
    If deducoesCompleta > descontoSimplificado Then
        CompararDeducaoCompletaSimplificada = "COMPLETA (Deduções: R$ " & _
            Format(deducoesCompleta, "#,##0.00") & ")"
    Else
        CompararDeducaoCompletaSimplificada = "SIMPLIFICADA (Desconto: R$ " & _
            Format(descontoSimplificado, "#,##0.00") & ")"
    End If
End Function

' ==============================================================================
' FUNÇÃO: VerificarLimiteEducacao
' DESCRIÇÃO: Verifica se despesa com educação está dentro do limite
' ==============================================================================
Public Function VerificarLimiteEducacao(valorDespesa As Double) As String
    If valorDespesa <= DEDUCAO_EDUCACAO_LIMITE Then
        VerificarLimiteEducacao = "OK - Dentro do limite"
    Else
        Dim excedente As Double
        excedente = valorDespesa - DEDUCAO_EDUCACAO_LIMITE
        VerificarLimiteEducacao = "⚠️ ATENÇÃO: Excede R$ " & Format(excedente, "#,##0.00")
    End If
End Function

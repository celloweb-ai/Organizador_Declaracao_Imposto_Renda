Attribute VB_Name = "ModuloProtecao"
' ==============================================================================
' MÓDULO DE PROTEÇÃO E BACKUP
' Organizador de Declaração de Imposto de Renda
' Desenvolvido por: Marcus Vasconcellos - DIO Bootcamp Santander 2025
' ==============================================================================

Option Explicit

' ==============================================================================
' SUBROTINA: ProtegerFormulas
' DESCRIÇÃO: Protege células com fórmulas, liberando apenas entrada de dados
' ==============================================================================
Public Sub ProtegerFormulas()
    Dim ws As Worksheet
    Dim celula As Range
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' Desprotege a planilha temporariamente
        ws.Unprotect
        
        ' Desbloqueia todas as células primeiro
        ws.Cells.Locked = False
        
        ' Bloqueia apenas células com fórmulas
        For Each celula In ws.UsedRange
            If celula.HasFormula Then
                celula.Locked = True
            End If
        Next celula
        
        ' Protege a planilha permitindo formatação
        ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                   AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                   AllowFormattingRows:=True
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "Fórmulas protegidas com sucesso!" & vbCrLf & _
           "Apenas campos de entrada podem ser editados.", _
           vbInformation, "LION APP - Proteção"
End Sub

' ==============================================================================
' SUBROTINA: DesprotegerTudo
' DESCRIÇÃO: Remove proteção de todas as planilhas
' ==============================================================================
Public Sub DesprotegerTudo()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
        ws.Cells.Locked = False
    Next ws
    
    MsgBox "Todas as planilhas desprotegidas!", vbInformation, "LION APP"
End Sub

' ==============================================================================
' SUBROTINA: CriarBackupAutomatico
' DESCRIÇÃO: Cria cópia de backup com timestamp
' ==============================================================================
Public Sub CriarBackupAutomatico()
    Dim caminhoOriginal As String
    Dim nomeArquivo As String
    Dim extensao As String
    Dim timestamp As String
    Dim caminhoBackup As String
    
    ' Gera timestamp no formato AAAAMMDD_HHMMSS
    timestamp = Format(Now, "YYYYMMDD_HHMMSS")
    
    ' Obtém caminho e nome do arquivo
    caminhoOriginal = ThisWorkbook.Path
    nomeArquivo = Replace(ThisWorkbook.Name, ".xlsm", "")
    nomeArquivo = Replace(nomeArquivo, ".xlsx", "")
    extensao = ".xlsx"
    
    ' Define caminho do backup
    caminhoBackup = caminhoOriginal & "\Backups\"
    
    ' Cria pasta Backups se não existir
    If Dir(caminhoBackup, vbDirectory) = "" Then
        MkDir caminhoBackup
    End If
    
    ' Salva cópia
    On Error Resume Next
    ThisWorkbook.SaveCopyAs caminhoBackup & nomeArquivo & "_BACKUP_" & timestamp & extensao
    
    If Err.Number = 0 Then
        MsgBox "Backup criado com sucesso!" & vbCrLf & _
               "Local: " & caminhoBackup & vbCrLf & _
               "Arquivo: " & nomeArquivo & "_BACKUP_" & timestamp & extensao, _
               vbInformation, "LION APP - Backup"
    Else
        MsgBox "Erro ao criar backup: " & Err.Description, vbCritical, "LION APP - Erro"
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' EVENTO: Workbook_BeforeClose
' DESCRIÇÃO: Cria backup automaticamente ao fechar (adicionar no ThisWorkbook)
' ==============================================================================
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     Dim resposta As VbMsgBoxResult
'     
'     resposta = MsgBox("Deseja criar um backup antes de fechar?", _
'                       vbYesNo + vbQuestion, "LION APP")
'     
'     If resposta = vbYes Then
'         Call ModuloProtecao.CriarBackupAutomatico
'     End If
' End Sub

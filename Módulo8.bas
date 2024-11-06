Attribute VB_Name = "M�dulo8"
Sub SALVAR_DIRETORIOS_INDIVIDUALMENTE()
    Dim docOrigem As Document
    Dim docDestino As Document
    Dim recordCount As Integer
    Dim i As Integer
    Dim pastaDestino As String
    Dim nomeArquivo As String
    Dim dialog As FileDialog
    Dim campoEmpresa As String
    Dim campo As MailMergeField
    
    ' Abra a janela de sele��o de pasta
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "Selecione a pasta onde os arquivos ser�o salvos"
    
    ' Se o usu�rio escolher uma pasta, continue
    If dialog.Show = -1 Then
        pastaDestino = dialog.SelectedItems(1) & "\"
    Else
        MsgBox "Nenhuma pasta foi selecionada. Opera��o cancelada.", vbExclamation
        Exit Sub
    End If
    
    ' Defina o documento de origem como o documento ativo
    Set docOrigem = ActiveDocument
    
    ' Obtenha o n�mero total de registros na mala direta
    recordCount = docOrigem.MailMerge.dataSource.recordCount
    
    ' Loop atrav�s de todos os registros da mala direta
    For i = 1 To recordCount
        ' Mover para o registro atual
        docOrigem.MailMerge.dataSource.ActiveRecord = i
        
        ' Mesclar para um novo documento apenas o registro atual
        docOrigem.MailMerge.Destination = wdSendToNewDocument
        docOrigem.MailMerge.dataSource.FirstRecord = i
        docOrigem.MailMerge.dataSource.LastRecord = i
        docOrigem.MailMerge.Execute
        
        ' Defina o documento mesclado como o documento de destino
        Set docDestino = ActiveDocument
        
        ' Tentar obter o valor do campo �Empresa�
        On Error Resume Next
        campoEmpresa = docOrigem.MailMerge.dataSource.DataFields("Empresa").Value
        On Error GoTo 0
        
        ' Verificar se o valor foi obtido corretamente
        If campoEmpresa = "" Then
            campoEmpresa = "Registro_" & i
        End If
        
        ' Remover caracteres inv�lidos para nome de arquivo
        campoEmpresa = Replace(campoEmpresa, "/", "_")
        campoEmpresa = Replace(campoEmpresa, "\", "_")
        campoEmpresa = Replace(campoEmpresa, ":", "_")
        campoEmpresa = Replace(campoEmpresa, "*", "_")
        campoEmpresa = Replace(campoEmpresa, "?", "_")
        campoEmpresa = Replace(campoEmpresa, """", "_")
        campoEmpresa = Replace(campoEmpresa, "<", "_")
        campoEmpresa = Replace(campoEmpresa, ">", "_")
        campoEmpresa = Replace(campoEmpresa, "|", "_")
        
        ' Defina o nome do arquivo com base no campo �Empresa�
        nomeArquivo = campoEmpresa & i & ".docx"
        
        ' Salvar o documento mesclado com a extens�o correta
        docDestino.SaveAs2 fileName:=pastaDestino & nomeArquivo, FileFormat:=wdFormatXMLDocument
        
        ' Fechar o documento mesclado
        docDestino.Close False
    Next i
    
    ' Notifica��o de conclus�o
    MsgBox "Todos os registros foram salvos na pasta selecionada.", vbInformation
End Sub


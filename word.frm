VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7584
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15312
   OleObjectBlob   =   "word.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ImportAndFormatTextFile()
    Dim dlg As FileDialog
    Dim txtFile As String
    Dim fileContent As String
    Dim fileLine As String
    Dim fontSize As Integer
    Dim lineText As String
    Dim para As Paragraph
    
    ' Criar um diálogo para selecionar o arquivo TXT
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Selecione um arquivo TXT"
        .Filters.Add "Text Files", "*.txt", 1
        .AllowMultiSelect = False
        
        If .Show <> -1 Then Exit Sub ' Cancelado pelo usuário
        txtFile = .SelectedItems(1)
    End With
    
    ' Abrir o arquivo para leitura
    Open txtFile For Input As #1
    
    ' Limpar o documento atual
    ActiveDocument.Content.Delete
    
    ' Ler o arquivo linha por linha
    Do While Not EOF(1)
        Line Input #1, fileLine
        
        ' Verificar se a linha começa com um número (tamanho da fonte)
        If IsNumeric(Split(fileLine, " ")(0)) Then
            fontSize = CInt(Split(fileLine, " ")(0))
            lineText = Mid(fileLine, InStr(fileLine, " ") + 1)
        Else
            fontSize = 12 ' Tamanho padrão se não houver número
            lineText = fileLine
        End If
        
        ' Adicionar o texto ao documento
        Set para = ActiveDocument.Content.Paragraphs.Add
        para.Range.Text = lineText
        para.Range.Font.Size = fontSize
        para.Range.InsertParagraphAfter
    Loop
    
    ' Fechar o arquivo
    Close #1
    
    MsgBox "Arquivo TXT importado e formatado com sucesso!", vbInformation
End Sub

Private Sub UserForm_Click()
ImportAndFormatTextFile
End Sub

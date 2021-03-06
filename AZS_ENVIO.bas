Attribute VB_Name = "AZS_ENVIO"
Option Explicit

Function EnviarOrcamentos(strEMAIL As String, strAssunto As String, strArquivo As String, strConteudo As String)
On Error GoTo EnviarEmail_err
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
    Dim x As Integer ' contador de linhas
    

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = strEMAIL
        .CC = ""
        .BCC = ""
        .Subject = strAssunto
        .Body = strConteudo
        .Attachments.Add strArquivo
        .Send
    End With
    On Error GoTo 0
    

EnviarEmail_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing
  
    Exit Function
EnviarEmail_err:
    MsgBox Err.Description
    Resume EnviarEmail_Fim
    
End Function

'Function EnviarEmail(strEMAIL As String, strAssunto As String, Optional intLinhaCaminhoArquivo As Integer, Optional intColunaCaminhoArquivo As Integer, Optional intArquivoControle As Integer)
'On Error GoTo EnviarEmail_err
'
'    Dim OutApp As Object
'    Dim OutMail As Object
'
'    Dim l As Integer, c As Integer ' L = LINHA | C = COLUNA
'    Dim x As Integer ' contador de linhas
'
'
'    Set OutApp = CreateObject("Outlook.Application")
'    Set OutMail = OutApp.CreateItem(0)
'
'    On Error Resume Next
'   ' Change the mail address and subject in the macro before you run it.
'    With OutMail
'        .To = strEMAIL
'        .CC = ""
'        .BCC = ""
'        .Subject = strAssunto 'ActiveSheet.Name
''        .Body = "Hello World!"
''        .Attachments.Add ActiveWorkbook.FullName
'        ' You can add other files by uncommenting the following line.
'
'        If intLinhaCaminhoArquivo > 0 And intColunaCaminhoArquivo > 0 And intArquivoControle > 0 Then
'
'            l = intLinhaCaminhoArquivo
'            c = intColunaCaminhoArquivo
'            For x = l To Range(ArquivoControle).Value
'                .Attachments.Add (Cells(l, c).Value)
'                l = l + 1
'            Next x
'
'        End If
'
'        ' In place of the following statement, you can use ".Display" to
'        ' display the mail.
'
'        .Send
'    End With
'    On Error GoTo 0
'
'
'EnviarEmail_Fim:
'    Set OutMail = Nothing
'    Set OutApp = Nothing
'
'    Exit Function
'EnviarEmail_err:
'    MsgBox Err.Description
'    Resume EnviarEmail_Fim
'
'End Function
'
'


Function EnviarEmail(strEMAIL As String, _
                     strAssunto As String, _
                     Anexo As Boolean, _
                     Optional strVendedor As String, _
                     Optional strControle As String, _
                     Optional BaseDeDados As String, _
                     Optional strConsulta As String, _
                     Optional strCampo As String)

On Error GoTo EnviarEmail_err
        
    Dim OutApp As Object
    Set OutApp = CreateObject("Outlook.Application")
    
    Dim OutMail As Object
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = strEMAIL
        .CC = ""
        .BCC = ""
        .Subject = strAssunto 'ActiveSheet.Name
'        .Body = "Hello World!"
'        .Attachments.Add ActiveWorkbook.FullName
        ' You can add other files by uncommenting the following line.
        
        If Anexo = True Then
        
            '   BASE DE DADOS
            Dim dbOrcamento As DAO.Database
            Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
            
            '   CONSULTA
            Dim rstOrcamentosAnexos As DAO.Recordset
            Set rstOrcamentosAnexos = dbOrcamento.OpenRecordset(strConsulta)
            
            While Not rstOrcamentosAnexos.EOF
                .Attachments.Add rstOrcamentosAnexos.Fields(strCampo)
                rstOrcamentosAnexos.MoveNext
            Wend
            
        
        End If


        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        
        .Send
    End With
    On Error GoTo 0
    

EnviarEmail_Fim:
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    If Anexo = True Then
        Set dbOrcamento = Nothing
        Set rstOrcamentosAnexos = Nothing
    End If
  
    Exit Function
EnviarEmail_err:
    MsgBox Err.Description
    Resume EnviarEmail_Fim
    
End Function




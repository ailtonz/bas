Attribute VB_Name = "basSynchronism"
Sub Listagem(strBanco As infBanco)
Dim connection As New ADODB.connection
Dim rstAtualizacoes As New ADODB.Recordset
Dim fld As ADODB.Field
Dim x As Integer: x = "13"

    Set connection = OpenConnection(strBanco)
    If connection.State = 1 Then
        Call rstAtualizacoes.Open("SELECT * FROM " & strBanco.strTabela & " WHERE vendedor = '" & strBanco.strVendedor & "' AND " & " controle = '" & strBanco.strControle & "'", connection, adOpenStatic, adLockOptimistic)
        
        '' CAMPOS
        For Each fld In rstAtualizacoes.Fields
            If fld.Name <> "Codigo" Then
                Range(strBanco.strFiltro.strCampo & x).Value = fld.Name
            End If
            x = x + 1
        Next
              
        '' VALORES
        x = "13"
        Do While Not rstAtualizacoes.EOF
            For Each fld In rstAtualizacoes.Fields
                If fld.Name <> "Codigo" Then
                    Range(strBanco.strFiltro.strValor & x).Value = rstAtualizacoes(fld.Name).Value
                End If
                x = x + 1
            Next
            rstAtualizacoes.MoveNext
        Loop
    
    Else
        MsgBox "Falha na conex�o com o banco de dados!", vbCritical + vbOKOnly, "Falha na conex�o com o banco."
    End If
    connection.Close

End Sub

Function Remessa(strOperacao As String, strLocal As infBanco, strServer As infBanco)
Dim connection As New ADODB.connection
Dim rstAtualizacoes As ADODB.Recordset
Dim strTabelas(3) As String, strSql As String, strUsuario As String, strControle As String
Set rstAtualizacoes = New ADODB.Recordset
strTabelas(1) = "Orcamentos"
strTabelas(2) = "OrcamentosAnexos"
strTabelas(3) = "OrcamentosCustos"

    ''Is Internet Connected
    If IsInternetConnected() = True Then
        Set connection = OpenConnection(strLocal)
        '' Is Connected
        If connection.State = 1 Then
            '' Tasks
            If strLocal.strSource = "Access2003" Then
                Call rstAtualizacoes.Open("SELECT DISTINCT controle, vendedor FROM OrcamentosAtualizacoes ", connection, adOpenStatic, adLockOptimistic)
            Else
                Call rstAtualizacoes.Open("SELECT DISTINCT controle, vendedor FROM OrcamentosAtualizacoes where usuario = '" & strLocal.strVendedor & "'", connection, adOpenStatic, adLockOptimistic)
            End If

            '' VERIFICAR ATUALIZA��ES
            Do While Not rstAtualizacoes.EOF

                '' ENVIAR/RECEBER DADOS
                For x = 1 To UBound(strTabelas)
                    strSql = "SELECT * FROM " & strTabelas(x) & " WHERE controle = '" & rstAtualizacoes.Fields("CONTROLE").Value & "' AND vendedor = '" & rstAtualizacoes.Fields("VENDEDOR").Value & "'"
                    EnvioDeDados strLocal, strServer, strSql
                Next x

                rstAtualizacoes.MoveNext
            Loop

            '' EXCLUIR ATUALIZA��ES LOCAIS
            delTasksADO strLocal
            
            MsgBox strOperacao & " ok!", vbInformation + vbOKOnly, strOperacao
        Else
            MsgBox "Falha na conex�o com o banco de dados!", vbCritical + vbOKOnly, "Falha na conex�o com o banco. (" & strOperacao & ")"
        End If
        connection.Close
    Else
        ' no connected
        MsgBox "SEM INTERNET.", vbOKOnly + vbExclamation
    End If
    
End Function

Sub EnvioDeDados(dbOrigem As infBanco, dbDestino As infBanco, strSql As String)

Dim Origem As New ADODB.connection
Set Origem = OpenConnection(dbOrigem)

Dim rstOrigem As ADODB.Recordset
Set rstOrigem = New ADODB.Recordset


Dim Destino As New ADODB.connection
Set Destino = OpenConnection(dbDestino)

Dim rstDestino As ADODB.Recordset
Set rstDestino = New ADODB.Recordset

Dim fld As ADODB.Field
Dim NewFile As Boolean: NewFile = False


    Call rstOrigem.Open(strSql, Origem, , adLockOptimistic)

    Call rstDestino.Open(strSql, Destino, adOpenDynamic, adLockOptimistic, adCmdText)


'    Destino.BeginTrans

    '' SE � EXISTE NO SERVER CADASTRAR
    If rstDestino.EOF Then
        NewFile = True
    End If

    Do While Not rstOrigem.EOF

        If NewFile Then
            rstDestino.AddNew
        End If

        For Each fld In rstDestino.Fields
            If fld.Name <> "Codigo" Then
                rstDestino(fld.Name).Value = rstOrigem(fld.Name).Value
            End If
        Next

        rstDestino.Update
        rstOrigem.MoveNext

    Loop

'    Destino.CommitTrans

    rstDestino.Close
    rstOrigem.Close

    Destino.Close
    Origem.Close

End Sub

Sub addTasksDAO(strBanco As infBanco)

Dim dbOrcamento As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSql As String

Set dbOrcamento = DBEngine.OpenDatabase(strBanco.strLocation & strBanco.strDatabase, False, False, "MS Access;PWD=" & strBanco.strPassword)
Set qdf = dbOrcamento.QueryDefs("admRemessa")

With qdf

    .Parameters("NM_VENDEDOR") = strBanco.strVendedor
    .Parameters("NM_CONTROLE") = strBanco.strControle
    
    .Execute
    
End With

qdf.Close
dbOrcamento.Close

End Sub

Sub addTasksADO(strBanco As infBanco, strBanco2 As infBanco)

Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)

Dim rst As ADODB.Recordset

Dim cd As ADODB.Command
Set cd = New ADODB.Command

With cd

    .ActiveConnection = connection

    .CommandText = "admOrcamentosAtualizacoesREMESSA"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, strBanco2.strVendedor)
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, strBanco2.strControle)


    Set rst = .Execute

End With

connection.Close

End Sub

Sub delTasksADO(strBanco As infBanco)

Dim connection As New ADODB.connection
Set connection = OpenConnection(strBanco)

Dim rst As ADODB.Recordset

Dim cmd As ADODB.Command
Set cmd = New ADODB.Command

With cmd

    .ActiveConnection = connection

    .CommandText = "admOrcamentosAtualizacoesEXCLUSAO"
    .CommandType = adCmdStoredProc
    .Parameters.Append cmd.CreateParameter("@NM_USUARIO", adVarChar, adParamInput, 50, strBanco.strVendedor)

    Set rst = .Execute

End With

connection.Close

End Sub

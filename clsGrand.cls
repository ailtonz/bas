VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Private pNome As String
'Private pProfissao As String
'Private pValorLiquido As Currency
'
'Public Property Get Profissao() As String
'    Profissao = pProfissao
'End Property
'
'Public Property Let Profissao(ByVal vProfissao As String)
'    pProfissao = vProfissao
'End Property
'
'Public Property Get Nome() As String
'    Nome = pNome
'End Property
'
'Public Property Let Nome(ByVal vNome As String)
'    pNome = vNome
'End Property
'
'Public Property Get ValorLiquido() As Currency
'    ValorLiquido = pValorLiquido
'End Property
'
'Public Property Let ValorLiquido(ByVal vValorLiquido As Currency)
'    pValorLiquido = vValorLiquido
'End Property
'
'Public Sub Insert(Orcamento As clsOrcamento, Grand As clsGrand)
'Dim Bnc As New clsBancos
'Dim connection As New ADODB.connection
'Set connection = Bnc.OpenConnection(Bnc)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'
'Set cd = New ADODB.Command
'With cd
'    .ActiveConnection = connection
'    .CommandText = "admNewGrand"
'    .CommandType = adCmdStoredProc
'    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
'    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
'
'    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
'    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
'    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)
'
'    Set rst = .Execute
'End With
'connection.Close
'
'End Sub
'
'Public Sub Update(banco As clsBancos, Orcamento As clsOrcamento, Grand As clsGrand)
'Dim Bnc As clsBancos
'Dim connection As New ADODB.connection
'Set connection = Bnc.OpenConnection(banco)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'
'Set cd = New ADODB.Command
'With cd
'    .ActiveConnection = connection
'    .CommandText = "admUpdGrand"
'    .CommandType = adCmdStoredProc
'    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
'    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
'
'    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
'    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
'    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)
'
'    Set rst = .Execute
'End With
'connection.Close
'
'End Sub
'
'Public Sub Delete(banco As clsBancos, Orcamento As clsOrcamento, Grand As clsGrand)
'Dim Bnc As clsBancos
'Dim connection As New ADODB.connection
'Set connection = Bnc.OpenConnection(banco)
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'
'Set cd = New ADODB.Command
'With cd
'    .ActiveConnection = connection
'    .CommandText = "admDelGrand"
'    .CommandType = adCmdStoredProc
'    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
'    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
'
'    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
'    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
'    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)
'
'    Set rst = .Execute
'End With
'connection.Close
'
'End Sub

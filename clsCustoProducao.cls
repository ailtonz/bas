VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustoProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCustoProducao As New Collection

Private pID As String
Private pPaginas As Integer
Private pValor As Currency
Private pTipo As String
Private pEstilo As String

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Paginas() As Integer
    Paginas = pPaginas
End Property

Public Property Let Paginas(ByVal vPaginas As Integer)
    pPaginas = vPaginas
End Property

Public Property Get Valor() As Currency
    Valor = pValor
End Property

Public Property Let Valor(ByVal vValor As Currency)
    pValor = vValor
End Property

Public Property Get Tipo() As String
    Tipo = pTipo
End Property

Public Property Let Tipo(ByVal vTipo As String)
    pTipo = vTipo
End Property

Public Property Get Estilo() As String
    Estilo = pEstilo
End Property

Public Property Let Estilo(ByVal vEstilo As String)
    pEstilo = vEstilo
End Property



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCustoProducao.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCustoProducao = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCustoProducao = Nothing
End Sub

Public Sub add(ByVal recCustoProducao As clsCustoProducao, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCustoProducao.add recCustoProducao, key, before, after
End Sub

Public Function count() As Integer
    count = pCustoProducao.count
End Function

Public Function remove(ByVal index As Integer)
    pCustoProducao.remove index
End Function

Public Function clsCustoProducao(ByVal index As Variant) As clsCustoProducao
    Set clsCustoProducao = pCustoProducao.Item(index)
End Function


Property Get Item(NameOrNumber As Variant) As clsCustoProducao
    'return this particular person
    Set Item = pCustoProducao(NameOrNumber)
End Property


Public Sub Insert(banco As clsBancos, Orcamento As clsOrcamentos, CustoProducao As clsCustoProducao)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admCustosProducaoNew"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, CustoProducao.Profissao)
    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, CustoProducao.Nome)
    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, CustoProducao.ValorLiquido)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Update(banco As clsBancos, Orcamento As clsOrcamentos, CustoProducao As clsCustoProducao)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admCustosProducaoUpd"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, CustoProducao.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Delete(banco As clsBancos,  CustoProducao As clsCustoProducao)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admCustosProducaoDel"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, CustoProducao.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

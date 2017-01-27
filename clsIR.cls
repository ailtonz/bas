VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIR As New Collection

Private pDescricao As String
Private pFaixaInicial As Currency
Private pFaixaFinal As Currency
Private pAliquota As Currency
Private pParcelaDeduzir As Currency

Public Property Get Descricao() As String
    Descricao = pDescricao
End Property

Public Property Let Descricao(ByVal vDescricao As String)
    pDescricao = vDescricao
End Property

Public Property Get FaixaInicial() As Currency
    FaixaInicial = pFaixaInicial
End Property

Public Property Let FaixaInicial(ByVal vFaixaInicial As Currency)
    pFaixaInicial = vFaixaInicial
End Property

Public Property Get FaixaFinal() As Currency
    FaixaFinal = pFaixaFinal
End Property

Public Property Let FaixaFinal(ByVal vFaixaFinal As Currency)
    pFaixaFinal = vFaixaFinal
End Property

Public Property Get Aliquota() As Currency
    Aliquota = pAliquota
End Property

Public Property Let Aliquota(ByVal vAliquota As Currency)
    pAliquota = vAliquota
End Property

Public Property Get ParcelaDeduzir() As Currency
    ParcelaDeduzir = pParcelaDeduzir
End Property

Public Property Let ParcelaDeduzir(ByVal vParcelaDeduzir As Currency)
    pParcelaDeduzir = vParcelaDeduzir
End Property



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pIR.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pIR = New Collection
End Sub

Private Sub Class_Terminate()
    Set pIR = Nothing
End Sub

Public Sub add(ByVal recIR As clsIR, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pIR.add recIR, key, before, after
End Sub

Public Function count() As Integer
    count = pIR.count
End Function

Public Function remove(ByVal index As Integer)
    pIR.remove index
End Function

Public Function clsIR(ByVal index As Variant) As clsIR
    Set clsIR = pIR.Item(index)
End Function


Property Get Item(NameOrNumber As Variant) As clsIR
    'return this particular person
    Set Item = pIR(NameOrNumber)
End Property


Public Sub Insert(banco As clsBancos, Orcamento As clsOrcamentos, IR As clsIR)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admIRNew"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)

    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, IR.Profissao)
    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, IR.Nome)
    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, IR.ValorLiquido)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Update(banco As clsBancos, Orcamento As clsOrcamentos, IR As clsIR)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admIRUpd"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)

    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, IR.Profissao)
    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, IR.Nome)
    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, IR.ValorLiquido)
    
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, IR.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Delete(banco As clsBancos, Orcamento As clsOrcamentos, IR As clsIR)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admIRDel"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, IR.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

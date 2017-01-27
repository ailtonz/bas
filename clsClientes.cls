VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pFK As String

Private pCnpjCpf As String
Private pIeRg As String
Private pNome As String
Private pNomeFantasia As String
Private pContatoNome As String
Private pContatoTelefone As String
Private pContatoEmail As String
Private pCadastroTipo As String
Private pCadastroPropaganda As String
Private pCadastroObservacao As String
Private pCadastroStatus As String
Private pEnderecoCep As String
Private pEnderecoNumero As String
Private pEnderecoComplemento As String
Private pEnderecoLogradouro As String
Private pEnderecoBairro As String
Private pEnderecoCidade As String
Private pEnderecoEstado As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vValor As String)
    pID = vValor
End Property

Public Property Get FK() As String
    FK = pFK
End Property

Public Property Let FK(ByVal vValor As String)
    pFK = vValor
End Property

Public Property Get CnpjCpf() As String
    CnpjCpf = pCnpjCpf
End Property

Public Property Let CnpjCpf(ByVal vValor As String)
    pCnpjCpf = vValor
End Property

Public Property Get IeRg() As String
    IeRg = pIeRg
End Property

Public Property Let IeRg(ByVal vValor As String)
    pIeRg = vValor
End Property

Public Property Get Nome() As String
    Nome = pNome
End Property

Public Property Let Nome(ByVal vValor As String)
    pNome = vValor
End Property

Public Property Get NomeFantasia() As String
    NomeFantasia = pNomeFantasia
End Property

Public Property Let NomeFantasia(ByVal vValor As String)
    pNomeFantasia = vValor
End Property

Public Property Get ContatoNome() As String
    ContatoNome = pContatoNome
End Property

Public Property Let ContatoNome(ByVal vValor As String)
    pContatoNome = vValor
End Property

Public Property Get ContatoTelefone() As String
    ContatoTelefone = pContatoTelefone
End Property

Public Property Let ContatoTelefone(ByVal vValor As String)
    pContatoTelefone = vValor
End Property

Public Property Get ContatoEmail() As String
    ContatoEmail = pContatoEmail
End Property

Public Property Let ContatoEmail(ByVal vValor As String)
    pContatoEmail = vValor
End Property

Public Property Get CadastroTipo() As String
    CadastroTipo = pCadastroTipo
End Property

Public Property Let CadastroTipo(ByVal vValor As String)
    pCadastroTipo = vValor
End Property

Public Property Get CadastroPropaganda() As String
    CadastroPropaganda = pCadastroPropaganda
End Property

Public Property Let CadastroPropaganda(ByVal vValor As String)
    pCadastroPropaganda = vValor
End Property

Public Property Get CadastroObservacao() As String
    CadastroObservacao = pCadastroObservacao
End Property

Public Property Let CadastroObservacao(ByVal vValor As String)
    pCadastroObservacao = vValor
End Property

Public Property Get CadastroStatus() As String
    CadastroStatus = pCadastroStatus
End Property

Public Property Let CadastroStatus(ByVal vValor As String)
    pCadastroStatus = vValor
End Property

Public Property Get EnderecoCep() As String
    EnderecoCep = pEnderecoCep
End Property

Public Property Let EnderecoCep(ByVal vValor As String)
    pEnderecoCep = vValor
End Property

Public Property Get EnderecoNumero() As String
    EnderecoNumero = pEnderecoNumero
End Property

Public Property Let EnderecoNumero(ByVal vValor As String)
    pEnderecoNumero = vValor
End Property

Public Property Get EnderecoComplemento() As String
    EnderecoComplemento = pEnderecoComplemento
End Property

Public Property Let EnderecoComplemento(ByVal vValor As String)
    pEnderecoComplemento = vValor
End Property

Public Property Get EnderecoLogradouro() As String
    EnderecoLogradouro = pEnderecoLogradouro
End Property

Public Property Let EnderecoLogradouro(ByVal vValor As String)
    pEnderecoLogradouro = vValor
End Property

Public Property Get EnderecoBairro() As String
    EnderecoBairro = pEnderecoBairro
End Property

Public Property Let EnderecoBairro(ByVal vValor As String)
    pEnderecoBairro = vValor
End Property

Public Property Get EnderecoCidade() As String
    EnderecoCidade = pEnderecoCidade
End Property

Public Property Let EnderecoCidade(ByVal vValor As String)
    pEnderecoCidade = vValor
End Property

Public Property Get EnderecoEstado() As String
    EnderecoEstado = pEnderecoEstado
End Property

Public Property Let EnderecoEstado(ByVal vValor As String)
    pEnderecoEstado = vValor
End Property

''---------------
'' administração
''---------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As clsClientes, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsClientes)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsClientes(ByVal i As Variant) As clsClientes
    Set clsClientes = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsClientes
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, obj As clsClientes) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "spClientes"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.ID)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)

    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)
    
    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 100, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.EnderecoCep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.EnderecoNumero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.EnderecoComplemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.EnderecoLogradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.EnderecoBairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.EnderecoCidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.EnderecoEstado)


    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Insert = False
  Else
    'Return a good return code
    Insert = True
  End If

Exit_Insert:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Insert:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()" & vbNewLine & Err.Description)
  Insert = False
  Resume Exit_Insert


End Function

Public Function Update(banco As clsBancos, obj As clsClientes) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "spClientes"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.ID)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)

    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)
    
    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 100, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.EnderecoCep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.EnderecoNumero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.EnderecoComplemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.EnderecoLogradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.EnderecoBairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.EnderecoCidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.EnderecoEstado)


    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Update = False
  Else
    'Return a good return code
    Update = True
  End If

Exit_Update:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Update:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()" & vbNewLine & Err.Description)
  Update = False
  Resume Exit_Update

End Function

Public Function Delete(banco As clsBancos, obj As clsClientes) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnectionNEW(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "spClientes"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@pID", adVarChar, adParamInput, 50, obj.ID)
    .Parameters.Append .CreateParameter("@pFK", adVarChar, adParamInput, 50, obj.FK)

    .Parameters.Append .CreateParameter("@pCnpjCpf", adVarChar, adParamInput, 14, obj.CnpjCpf)
    .Parameters.Append .CreateParameter("@pIeRg", adVarChar, adParamInput, 16, obj.IeRg)
    .Parameters.Append .CreateParameter("@pNome", adVarChar, adParamInput, 100, obj.Nome)
    .Parameters.Append .CreateParameter("@pNomeFantasia", adVarChar, adParamInput, 50, obj.NomeFantasia)
    
    .Parameters.Append .CreateParameter("@pContatoNome", adVarChar, adParamInput, 50, obj.ContatoNome)
    .Parameters.Append .CreateParameter("@pContatoTelefone", adVarChar, adParamInput, 50, obj.ContatoTelefone)
    .Parameters.Append .CreateParameter("@pContatoEmail", adVarChar, adParamInput, 120, obj.ContatoEmail)
    
    .Parameters.Append .CreateParameter("@pCadastroTipo", adVarChar, adParamInput, 2, obj.CadastroTipo)
    .Parameters.Append .CreateParameter("@pCadastroPropaganda", adVarChar, adParamInput, 50, obj.CadastroPropaganda)
    .Parameters.Append .CreateParameter("@pCadastroStatus", adVarChar, adParamInput, 10, obj.CadastroStatus)
    .Parameters.Append .CreateParameter("@pCadastroObservacao", adVarChar, adParamInput, 100, obj.CadastroObservacao)
    
    .Parameters.Append .CreateParameter("@pEnderecoCep", adVarChar, adParamInput, 8, obj.EnderecoCep)
    .Parameters.Append .CreateParameter("@pEnderecoNumero", adVarChar, adParamInput, 10, obj.EnderecoNumero)
    .Parameters.Append .CreateParameter("@pEnderecoComplemento", adVarChar, adParamInput, 50, obj.EnderecoComplemento)
    .Parameters.Append .CreateParameter("@pEnderecoLogradouro", adVarChar, adParamInput, 200, obj.EnderecoLogradouro)
    .Parameters.Append .CreateParameter("@pEnderecoBairro", adVarChar, adParamInput, 50, obj.EnderecoBairro)
    .Parameters.Append .CreateParameter("@pEnderecoCidade", adVarChar, adParamInput, 100, obj.EnderecoCidade)
    .Parameters.Append .CreateParameter("@pEnderecoEstado", adVarChar, adParamInput, 2, obj.EnderecoEstado)


    Set rst = .Execute(lRecordsAffected)
End With
Connection.Close

  If lRecordsAffected = 0 Then
    Delete = False
  Else
    'Return a good return code
    Delete = True
  End If

Exit_Delete:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Delete:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()" & vbNewLine & Err.Description)
  Delete = False
  Resume Exit_Delete

End Function

Public Function getClientes(banco As clsBancos) As clsClientes
On Error GoTo Err_getClientes
Dim colObj As New clsClientes
Dim recObj As New clsClientes

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from vw_Clientes"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsClientes
    With recObj
       
        .ID = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .CnpjCpf = CStr(IIf(IsNull(rst.Fields("CnpjCpf")), "", rst.Fields("CnpjCpf")))
        .IeRg = CStr(IIf(IsNull(rst.Fields("IeRg")), "", rst.Fields("IeRg")))
        .Nome = CStr(IIf(IsNull(rst.Fields("Nome")), "", rst.Fields("Nome")))
        .NomeFantasia = CStr(IIf(IsNull(rst.Fields("NomeFantasia")), "", rst.Fields("NomeFantasia")))
        
        .ContatoNome = CStr(IIf(IsNull(rst.Fields("ContatoNome")), "", rst.Fields("ContatoNome")))
        .ContatoTelefone = CStr(IIf(IsNull(rst.Fields("ContatoTelefone")), "", rst.Fields("ContatoTelefone")))
        .ContatoEmail = CStr(IIf(IsNull(rst.Fields("ContatoEmail")), "", rst.Fields("ContatoEmail")))
        
        .CadastroTipo = CStr(IIf(IsNull(rst.Fields("CadastroTipo")), "", rst.Fields("CadastroTipo")))
        .CadastroPropaganda = CStr(IIf(IsNull(rst.Fields("CadastroPropaganda")), "", rst.Fields("CadastroPropaganda")))
        .CadastroStatus = CStr(IIf(IsNull(rst.Fields("CadastroStatus")), "", rst.Fields("CadastroStatus")))
        .CadastroObservacao = CStr(IIf(IsNull(rst.Fields("CadastroObservacao")), "", rst.Fields("CadastroObservacao")))
        
        .EnderecoCep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .EnderecoNumero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .EnderecoComplemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .EnderecoLogradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .EnderecoBairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .EnderecoCidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .EnderecoEstado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getClientes:
    Set getClientes = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getClientes:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getClientes()" & vbNewLine & Err.Description)
    Resume Exit_getClientes

End Function

Public Function getClientesID(banco As clsBancos, i As String) As clsClientes
On Error GoTo Err_getClientesID
Dim colObj As New clsClientes
Dim recObj As New clsClientes

Dim cnn As New ADODB.Connection
Set cnn = OpenConnectionNEW(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from vw_Clientes where id = " & i & ""
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsClientes
    With recObj
       
        .ID = CStr(IIf(IsNull(rst.Fields("ID")), "", rst.Fields("ID")))
        .FK = CStr(IIf(IsNull(rst.Fields("FK")), "", rst.Fields("FK")))
        
        .CnpjCpf = CStr(IIf(IsNull(rst.Fields("CnpjCpf")), "", rst.Fields("CnpjCpf")))
        .IeRg = CStr(IIf(IsNull(rst.Fields("IeRg")), "", rst.Fields("IeRg")))
        .Nome = CStr(IIf(IsNull(rst.Fields("Nome")), "", rst.Fields("Nome")))
        .NomeFantasia = CStr(IIf(IsNull(rst.Fields("NomeFantasia")), "", rst.Fields("NomeFantasia")))
        
        .ContatoNome = CStr(IIf(IsNull(rst.Fields("ContatoNome")), "", rst.Fields("ContatoNome")))
        .ContatoTelefone = CStr(IIf(IsNull(rst.Fields("ContatoTelefone")), "", rst.Fields("ContatoTelefone")))
        .ContatoEmail = CStr(IIf(IsNull(rst.Fields("ContatoEmail")), "", rst.Fields("ContatoEmail")))
        
        .CadastroTipo = CStr(IIf(IsNull(rst.Fields("CadastroTipo")), "", rst.Fields("CadastroTipo")))
        .CadastroPropaganda = CStr(IIf(IsNull(rst.Fields("CadastroPropaganda")), "", rst.Fields("CadastroPropaganda")))
        .CadastroStatus = CStr(IIf(IsNull(rst.Fields("CadastroStatus")), "", rst.Fields("CadastroStatus")))
        .CadastroObservacao = CStr(IIf(IsNull(rst.Fields("CadastroObservacao")), "", rst.Fields("CadastroObservacao")))
        
        .EnderecoCep = CStr(IIf(IsNull(rst.Fields("EnderecoCep")), "", rst.Fields("EnderecoCep")))
        .EnderecoNumero = CStr(IIf(IsNull(rst.Fields("EnderecoNumero")), "", rst.Fields("EnderecoNumero")))
        .EnderecoComplemento = CStr(IIf(IsNull(rst.Fields("EnderecoComplemento")), "", rst.Fields("EnderecoComplemento")))
        .EnderecoLogradouro = CStr(IIf(IsNull(rst.Fields("EnderecoLogradouro")), "", rst.Fields("EnderecoLogradouro")))
        .EnderecoBairro = CStr(IIf(IsNull(rst.Fields("EnderecoBairro")), "", rst.Fields("EnderecoBairro")))
        .EnderecoCidade = CStr(IIf(IsNull(rst.Fields("EnderecoCidade")), "", rst.Fields("EnderecoCidade")))
        .EnderecoEstado = CStr(IIf(IsNull(rst.Fields("EnderecoEstado")), "", rst.Fields("EnderecoEstado")))
        
        colObj.all recObj
    End With
    rst.MoveNext
Loop


Exit_getClientesID:
    Set getClientesID = colObj
    Set recObj = Nothing
    Set colObj = Nothing

    Exit Function

Err_getClientesID:
    Call MsgBox("Class: " & TypeName(Me) & " , Function: getClientesID()" & vbNewLine & Err.Description)
    Resume Exit_getClientesID

End Function

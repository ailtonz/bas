VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As Collection

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

Public Sub add(ByVal recGrand As clsGrand, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add recGrand, key, before, after
End Sub

Public Function count() As Integer
    count = pCol.count
End Function

Public Function remove(ByVal index As Integer)
    pCol.remove index
End Function

Public Function clsGrand(ByVal index As Variant) As clsGrand
    Set clsGrand = pCol.item(index)
End Function

Public Sub Insert(Banco As clsBanco, Orcamento As clsOrcamento, Grand As clsGrand)
Dim connection As New ADODB.connection
Set connection = OpenConnection(Banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = connection
    .CommandText = ""
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
    .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)

    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
    .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
    .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)

    Set rst = .Execute
End With
connection.Close

End Sub


Public Sub mostrar(ByVal clsGrand As clsGrand)

    MsgBox clsGrand.Nome
    MsgBox clsGrand.Profissao
    MsgBox clsGrand.ValorLiquido
    
End Sub

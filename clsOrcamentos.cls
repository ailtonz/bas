VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrcamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pEstilo As String

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Estilo() As String
    Estilo = pEstilo
End Property

Public Property Let Estilo(ByVal vEstilo As String)
    pEstilo = vEstilo
End Property


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

Public Sub add(ByVal rec As clsOrcamentos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Function count() As Integer
    count = pCol.count
End Function

Public Function remove(ByVal index As Integer)
    pCol.remove index
End Function

Public Function clsOrcamentos(ByVal index As Variant) As clsOrcamentos
    Set clsOrcamentos = pCol.Item(index)
End Function


Property Get Item(NameOrNumber As Variant) As clsOrcamentos
    'return this particular person
    Set Item = pCol(NameOrNumber)
End Property


Public Sub Insert(banco As clsBancos, obj As clsOrcamentos)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admEstiloNew"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Update(banco As clsBancos, obj As clsOrcamentos)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admEstiloUpd"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_ESTILO", adVarChar, adParamInput, 50, obj.Estilo)
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

Public Sub Delete(banco As clsBancos, obj As clsOrcamentos)
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admEstiloDel"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)

    Set rst = .Execute
End With
Connection.Close

End Sub

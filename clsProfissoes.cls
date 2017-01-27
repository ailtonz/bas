VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProfissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pProfissao As String

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Profissao() As String
    Profissao = pProfissao
End Property

Public Property Let Profissao(ByVal vProfissao As String)
    pProfissao = vProfissao
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

Public Sub add(ByVal rec As clsProfissoes, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub allProfissoes(ByVal rec As clsProfissoes)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Integer
    count = pCol.count
End Function

Public Function remove(ByVal index As Integer)
    pCol.remove index
End Function

Public Function clsProfissoes(ByVal index As Variant) As clsProfissoes
    Set clsProfissoes = pCol.Item(index)
End Function

Property Set Itens(c As clsProfissoes)
    Set pCol = c
End Property

Property Get Item(NameOrNumber As Variant) As clsProfissoes
    'return this particular person
    Set Item = pCol(NameOrNumber)
End Property

Public Function Insert(banco As clsBancos, obj As clsProfissoes) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admProfissaoNew"
    .CommandType = adCmdStoredProc

    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, obj.Profissao)

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
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()")
  Insert = False
  Resume Exit_Insert


End Function

Public Function Update(banco As clsBancos, obj As clsProfissoes) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admProfissaoUpd"
    .CommandType = adCmdStoredProc
    .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, obj.Profissao)
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)

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
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()")
  Update = False
  Resume Exit_Update

End Function

Public Function Delete(banco As clsBancos, obj As clsProfissoes) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

Set cd = New ADODB.Command
With cd
    .ActiveConnection = Connection
    .CommandText = "admProfissaoDel"
    .CommandType = adCmdStoredProc
    
    .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, obj.ID)

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
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()")
  Delete = False
  Resume Exit_Delete

End Function

'Public Function getProfissoes(banco As clsBancos) As ADODB.Recordset
'Dim cnn As New ADODB.Connection
'Set cnn = OpenConnection(banco)
'
'Dim cd As ADODB.Command
'Set cd = New ADODB.Command
'
'Dim rst As ADODB.Recordset
'
'With cd
'    .ActiveConnection = cnn
'    .CommandText = "select * from qryProfissoes"
'    .CommandType = adCmdText
'    Set rst = .Execute
'End With
'
'Set getProfissoes = rst
'
'cnn.Close
'
'Set rst = Nothing
'Set cd = Nothing
'
'End Function


'Public Function getProfissoes(banco As clsBancos) As clsProfissoes
'Dim obj As clsProfissoes
'
'Dim cnn As New ADODB.Connection
'Set cnn = OpenConnection(banco)
'
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'
'Set cd = New ADODB.Command
'With cd
'    .ActiveConnection = cnn
'    .CommandText = "select * from qryProfissoes"
'    .CommandType = adCmdText
'    Set rst = .Execute
'End With
'
'Set obj = New clsProfissoes
'Do While Not rst.EOF
'
'    With obj
'        .ID = rst.Fields("codCategoria")
'        .Profissao = rst.Fields("Descricao")
'        .add obj
'    End With
'
'    rst.MoveNext
'Loop
'
'Set getProfissoes = AddProfissao(obj)
'
'cnn.Close
'
'Set obj = Nothing
'Set rst = Nothing
'Set cnn = Nothing
'
'End Function



Public Function getProfissao(banco As clsBancos) As clsProfissoes
Dim obj As clsProfissoes
Set obj = New clsProfissoes

Dim cnn As New ADODB.Connection
Set cnn = OpenConnection(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryProfissoes"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    With obj
        .ID = rst.Fields("codCategoria")
        .Profissao = rst.Fields("Descricao")
        .add obj
    End With
    rst.MoveNext
Loop

Set getProfissao = obj

Set obj = Nothing

End Function




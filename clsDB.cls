VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pBancos As New Collection

Private pSource As String
Private pDriver As String
Private pLocation As String
Private pDatabase As String
Private pUser As String
Private pPassword As String
Private pPort As String
Private pTabela As String

Public Property Get Source() As String
    Source = pSource
End Property

Public Property Let Source(ByVal vSource As String)
    pSource = vSource
End Property

Public Property Get Driver() As String
    Driver = pDriver
End Property

Public Property Let Driver(ByVal vDriver As String)
    pDriver = vDriver
End Property

Public Property Get Location() As String
    Location = pLocation
End Property

Public Property Let Location(ByVal vLocation As String)
    pLocation = vLocation
End Property

Public Property Get Database() As String
    Database = pDatabase
End Property

Public Property Let Database(ByVal vDatabase As String)
    pDatabase = vDatabase
End Property

Public Property Get User() As String
    User = pUser
End Property

Public Property Let User(ByVal vUser As String)
    pUser = vUser
End Property

Public Property Get Password() As String
    Password = pPassword
End Property

Public Property Let Password(ByVal vPassword As String)
    pPassword = vPassword
End Property

Public Property Get Port() As String
    Port = pPort
End Property

Public Property Let Port(ByVal vPort As String)
    pPort = vPort
End Property

Public Property Get Tabela() As String
    Tabela = pTabela
End Property

Public Property Let Tabela(ByVal vTabela As String)
    pTabela = vTabela
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pBancos.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pBancos = New Collection
End Sub

Private Sub Class_Terminate()
    Set pBancos = Nothing
End Sub

Public Sub add(ByVal recBanco As clsDB, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pBancos.add recBanco, key, before, after
End Sub

Public Function count() As Integer
    count = pBancos.count
End Function

Public Function remove(ByVal index As Integer)
    pBancos.remove index
End Function

Public Function clsBanco(ByVal index As Variant) As clsDB
    Set clsBanco = pBancos.Item(index)
End Function


Public Function OpenConnection() As ADODB.Connection
'' Build the connection string depending on the source
Dim connectionString As String
    
Select Case pSource
    Case "Access"
        connectionString = "Provider=" & pDriver & ";Data Source=" & pDatabase
    Case "Access2003"
        connectionString = "Driver={" & pDriver & "};Dbq=" & pLocation & pDatabase & ";Uid=" & pUser & ";PWD=" & pPassword & ""
    Case "SQLite"
        connectionString = "Driver={" & pDriver & "};Database=" & pDatabase
    Case "MySQL"
        connectionString = "Driver={" & pDriver & "};Server=" & pLocation & ";Database=" & pDatabase & ";PORT=" & pPort & ";UID=" & pUser & ";PWD=" & pPassword
    Case "PostgreSQL"
        connectionString = "Driver={" & pDriver & "};Server=" & pLocation & ";Database=" & pDatabase & ";UID=" & pUser & ";PWD=" & pPassword
End Select

'' Create and open a new connection to the selected source
Set OpenConnection = New ADODB.Connection
Call OpenConnection.Open(connectionString)
   
End Function

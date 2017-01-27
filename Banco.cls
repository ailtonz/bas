VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_TENTATIVAS        As Integer = 3
Private Const MAX_SLEEP             As Long = 500
Private iSql                        As String
Private iCn                         As New ADODB.Connection
Private iRs                         As New ADODB.Recordset

Property Get sql() As String
    sql = iSql
End Property
Property Let sql(Value As String)
    iSql = Value
End Property
Property Get rs() As ADODB.Recordset
    Set rs = iRs
End Property
Property Let ConnectionString(Value As String)
    iCn.ConnectionString = Value
End Property
'================================================================================================================
'=   FIM DAS PROPRIEDADES DA CLASSE   ===========================================================================
'================================================================================================================
Public Function Sel(Optional sql As String, Optional pManterConexao As Boolean) As Boolean
    Dim intTentativas As Integer
    Dim strDesc As String
    If sql <> vbNullString Then iSql = sql
    If iSql = vbNullString Then Exit Function
    
    On Error Resume Next
        If conecta Then
        
            iRs.Close
            Set iRs = Nothing
            iRs.CursorLocation = adUseClient
    
            Do Until intTentativas = MAX_TENTATIVAS
                err.Clear
                iRs.Open iSql, iCn, adOpenDynamic, adLockPessimistic
                'iRs.Open iSql, iCn, ADODB.adOpenForwardOnly, adLockReadOnly
                If err.Number = 0 Then
                    intTentativas = MAX_TENTATIVAS
                    Set iRs.ActiveConnection = Nothing
                    fecha
                    Sel = True
                Else
                    strDesc = err.Description
                    intTentativas = intTentativas + 1
                    Sleep MAX_SLEEP
                End If
            Loop
        Else
            RaiseError "Banco.Conecta", "Ocorreu uma falha ao tentar estabelecer uma conexão com a base de dados." & vbNewLine & strDesc
        End If
    On Error GoTo 0
    If Not Sel Then
        'Debug.Print " -- Erro " & Now() & " -- " & vbNewLine & iSql
        'RaiseError "Banco.Sel", "Ocorreu uma falha ao tentar realizar a seguinte consulta na base de dados:" & vbNewLine & "   " & iSql & vbNewLine & "   " & strDesc
    End If
End Function

Public Function Exe(Optional sql As String, Optional pManterConexao As Boolean) As Boolean
    Dim intTentativas As Integer
    If sql <> vbNullString Then iSql = sql
    If iSql = vbNullString Then Exit Function
    On Error Resume Next
        If conecta Then
            Do Until intTentativas = MAX_TENTATIVAS
                err.Clear
                iCn.Execute iSql, adExecuteNoRecords
                If err.Number = 0 Then
                    intTentativas = MAX_TENTATIVAS
                    fecha
                    Exe = True
                Else
                    intTentativas = intTentativas + 1
                    Sleep MAX_SLEEP
                End If
            Loop
        Else
            RaiseError "Banco.Conecta", "Ocorreu uma falha ao tentar estabelecer uma conexão com a base de dados."
        End If
    On Error GoTo 0
    If Not Exe Then
        'Debug.Print " -- Erro " & Now() & " -- " & vbNewLine & iSql
        RaiseError "Banco.Exe", "Ocorreu uma falha ao tentar realizar a seguinte instrução na base de dados:" & vbNewLine & "   " & iSql & vbNewLine
    End If
End Function
Public Function conecta() As Boolean
    Dim intTentativas As Integer
    If iCn.State = 0 Then
        Do Until intTentativas = MAX_TENTATIVAS
            err.Clear
            iCn.ConnectionTimeout = 3600
            iCn.CommandTimeout = 3600
            iCn.Open
            If err.Number = 0 Then
                intTentativas = MAX_TENTATIVAS
                conecta = True
            Else
                intTentativas = intTentativas + 1
                Sleep MAX_SLEEP
            End If
        Loop
    Else
        conecta = True
    End If
    Set iRs = New ADODB.Recordset
End Function
Public Function fecha()
    On Error Resume Next
        iCn.Close
    On Error GoTo 0
End Function
Private Sub Class_Initialize()
    'iCn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Persist Security Info=true;Data Source=" & ThisWorkbook.Path & "\BrasilPrev.mdb;Jet OLEDB:Database Password=0cdoci5f!"
    iCn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & shtMenu.txtDatabasePath & ";Extended Properties=Excel 12.0"
End Sub
Public Function possuiAcesso() As Boolean
    ThisWorkbook.Application.ScreenUpdating = False
    ThisWorkbook.Application.DisplayAlerts = False
    On Error Resume Next
    Open ThisWorkbook.Path & "\" & Environ("UserName") & "TEMP.clm" For Output As #1
    Close #1
    If err.Number <> 0 Then
        MsgBox "O usuário " & Environ("UserName") & " não possui acesso para gravação na pasta " & ThisWorkbook.Path & ".", vbCritical, TITULO_MSG
        possuiAcesso = False
    Else
        possuiAcesso = True
    End If
    Kill ThisWorkbook.Path & "\*.clm"
    On Error GoTo 0
    ThisWorkbook.Application.ScreenUpdating = True
    ThisWorkbook.Application.DisplayAlerts = True
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPw3270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function hllapi_init Lib "libhllapi.dll" (ByVal tp As String) As Long
Private Declare Function hllapi_deinit Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_get_revision Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_connect Lib "libhllapi.dll" (ByVal uri As String, ByVal wait As Integer) As Long
Private Declare Function hllapi_disconnect Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_wait_for_ready Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
Private Declare Function hllapi_get_screen_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare Function hllapi_enter Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_get_message_id Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_set_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare Function hllapi_wait Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
Private Declare Function hllapi_pfkey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare Function hllapi_pakey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare Function hllapi_cmp_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare Function hllapi_is_connected Lib "libhllapi.dll" () As Long
Private Declare Function hllapi_erase Lib "libhllapi.dll" () As Long


Public Sub INICIO(sServidor As String, sUser As String, sPwd As String)
Dim iTentativaLogin As Integer

If iniciaEmulador Then
   If ConectarServidor(sServidor) Then
      For iTentativaLogin = 1 To 5
          If loginCSO(sUser, sPwd, sServidor) Then
             Exit For
          End If
      Next iTentativaLogin
   End If
End If

End Sub

Public Sub TERMINIO()
    Call hllapi_disconnect
    Call hllapi_deinit
End Sub

Public Sub CONTAS_PENDENTES(sPath As String)
    Call hllapi_pakey(1)
    Call hllapi_wait_for_ready(1)
    gravaLog "CONTAS_PENDENTES", sPath
End Sub

Public Sub VOLTAR_CADASTRO(sPath As String)
    Call hllapi_enter
    Call hllapi_wait_for_ready(3)
    gravaLog "VOLTAR_CADASTRO", srtPath
End Sub

Public Sub EMITIR_COR(sPath As String)
    Call hllapi_wait_for_ready(3)
    Call hllapi_pfkey(6)
    Call hllapi_wait_for_ready(3)
    gravaLog "EMITIR_COR", srtPath
End Sub

Public Sub PAUSA()
    Call hllapi_wait_for_ready(3)
End Sub

Public Sub ENTRAR_DADOS(sPath As String)
    Call hllapi_enter
    Call hllapi_enter
    Call hllapi_wait_for_ready(3)
    gravaLog "ENTRAR_DADOS", srtPath
End Sub

Public Sub LIMPAR_DADOS(sPath As String)
    Call hllapi_erase
    Call hllapi_wait_for_ready(3)
    gravaLog "LIMPAR_DADOS", srtPath
End Sub


Public Function iniciaEmulador() As Boolean
'==========================================
'= Função inicia a emulação do PW3270
'==========================================
On Error GoTo errConnect

Dim iConectEmulador As Integer
iniciaEmulador = False

'status emulador = 9, não iniciado
'status emulador = 0, iniciado
iConectEmulador = hllapi_is_connected
If iConectEmulador = 9 Then
    'trecho de chamada para iniciacao do emulador
    iConectEmulador = hllapi_init("")
    If iConectEmulador = 0 Then iniciaEmulador = True
    Exit Function
ElseIf iConectEmulador = 0 Then
    iniciaEmulador = True
    Exit Function
End If

Exit Function
errConnect:
    iniciaEmulador = False
End Function

Public Function ConectarServidor(sServidor As String) As Boolean
'============================================
'= Função Realiza a Conexao com o servidor
'============================================
On Error GoTo errConectar

    Dim iConectServidor As Integer
    Dim iIndex          As Integer
    
    Call hllapi_disconnect
    Conectar = False
       
    'Enquanto a conexão com o servidor for igual a 0
    'Realiza a tentativa de Conexão
    '===============================================
    iConectServidor = 0
    For iIndex = 1 To 5
        If iConectServidor = 0 Then
            iConectServidor = hllapi_connect(sServidor, 500)
            Call hllapi_wait_for_ready(5)
        Else
            ConectarServidor = True
            Exit Function
        End If
    Next iIndex
        
    Exit Function
errConectar:
    Conectar = False
End Function



Public Sub LogOut(iRow As Integer, iCol As Integer)
    Call hllapi_set_text_at(iRow, iCol, "/F")
    Call hllapi_enter
    Call hllapi_wait_for_ready(1)
End Sub

Public Sub IntoText(iRow As Integer, iCol As Integer, sTexto As String, iWait As Integer, bEnter As Boolean)
    Call hllapi_set_text_at(iRow, iCol, sTexto)
    Call hllapi_wait_for_ready(iWait)
        
    If bEnter = True Then
        Call hllapi_enter
        Call hllapi_wait_for_ready(iWait)
    End If
End Sub

Public Function RetvTela() As Variant
'======================================================================
'= Função Retorna a Posição da tela do CSO conforme parametros informado
'======================================================================

    Dim sTexto As String
    Dim iIndex As Integer
    Dim vTela(24) As Variant
    Dim iValida As Integer
    
    On Error GoTo errHandler
    sTexto = Space(2000)
    
    Call hllapi_get_screen_at(1, 1, sTexto)
    iValida = 1
    For iIndex = 1 To 24
        vTela(iIndex) = Mid(sTexto, iValida, 80) & vbCr
        iValida = iValida + 80
    Next iIndex
    
    RetvTela = vTela
    Exit Function

errHandler:
    RetvTela = Space(0)
End Function

Public Function RetTextoTela(iRow As Integer, iColuna As Integer, iTamanho) As String
'======================================================================
'= Função Retorna a Posição da tela do CSO conforme parametros informado
'======================================================================

    Dim sTexto As String
    Dim iIndex As Integer
    
    On Error GoTo errHandler
    sTexto = Space(iTamanho)
    Call hllapi_get_screen_at(iRow, iColuna, sTexto)
        
    RetTextoTela = sTexto
    Exit Function

errHandler:
    RetTextoTela = Space(0)
End Function

Public Function CopiaTela(sProcesso As String) As Variant
    
    Dim iIndex      As Integer
    Dim sSepara     As String
    Dim aList(28)   As Variant
    Dim vTela       As Variant
            
    vTela = RetvTela
    
    For iIndex = 1 To 80
        sSepara = sSepara & "-"
    Next iIndex
    
    aList(1) = sSepara
    aList(2) = sProcesso
       
    For iIndex = 1 To 24
        aList(iIndex + 2) = vTela(iIndex)
    Next iIndex
    aList(28) = sSepara & vbCrLf
    
    CopiaTela = aList
    
End Function

Public Sub gravaLog(sLog As String, sPath As String)
    Dim sFile As String
    Dim vTela As Variant
        
    If bLog Then
        sFile = sPath & "\Log_" & Replace(Date, "/", "_") & ".txt"
        If createTxtLog(sFile) Then
            vTela = CopiaTela(sLog)
            PrintLog sFile, vTela
        End If
    End If
        
End Sub

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Public Function createTxtLog(strFile As String) As Boolean
'=========================================================
'Função Cria Arquivo Txt para Log das telas do CSO
'=========================================================

    On Error GoTo errHandler
    
    Dim fso     As Object
    Dim iIndex  As Integer
    Dim oFile   As Object
    
    With Me
        
        'Valida se já existe o arquivo na pasta Raiz
        '===========================================
        If .FileExists(strFile) Then
            createTxtLog = True
            Exit Function
        Else
            
            'Se nao existir o arquivo, é criado
            '==================================
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set oFile = fso.CreateTextFile(strFile)
            
            If .FileExists(strFile) Then
                createTxtLog = True
            Else
                createTxtLog = False
            End If
            
            oFile.Close
            Set fso = Nothing
            Set oFile = Nothing
        End If
    
    End With
    
    Exit Function
    
errHandler:
    createTxtLog = False
End Function

Public Sub PrintLog(sFile As String, arrayTela As Variant)
    
    Dim iIndex As Integer
    
    Open sFile For Append As #1
    For iIndex = 1 To UBound(arrayTela)
        Write #1, arrayTela(iIndex)
    Next iIndex
    
    Close #1
End Sub


Public Function loginCSO(strUser As String, strPwd As String, sServidor As String) As Boolean
    
'=====================================================================================
'Função Recebe usuário e senha e acessa o emulador Pw3270 criando uma emolução oculta
'=====================================================================================
   
On Error GoTo errHandler
   
    Dim iConnect    As Long
    Dim sTela       As String
    Dim sOp         As String
        
    'Valida se a Tela Atual é a tela de Login
    '========================================
    sTela = RetTextoTela(18, 5, 7)
    If sTela = "Cliente" Then
        
        'Seta usuário e senha nas posicoes
        '========================================================
        Call IntoText(18, 15, Left(strUser, 7), 1, False) 'Usuário
        Call IntoText(19, 15, strPwd, 5, True)            'Senha
        '=========================================================
        
        'Valida se o login foi realizado
        '==========================================================
        
        'Verifica se é a tela inicial após o login
        sTela = RetTextoTela(5, 6, 7)
        
        'Verifica se existe processo preso na seção
        sOp = Trim(RetTextoTela(7, 1, 80))
    
         
        If sTela = "Monitor" And Len(sOp) = 0 Then 'True para logon sem seção presa
            loginCSO = True
            Exit Function
        ElseIf sTela = "Monitor" And Len(sOp) > 0 Then 'False e LogOut para logon com seção presa
            Call LogOut(23, 29)
            loginCSO = False
            Exit Function
        Else
            ConectarServidor sServidor
            loginCSO = False
        End If
    End If
    
    Exit Function

errHandler:
    loginCSO = False
End Function

Public Function acessoOperacao(sPasso1 As String, sPasso2 As String, _
                              sValida As String, iValida As Integer, sPath As String) As Boolean

'==============================================================================
'Variável sPasso1 para a entrada da 1 tela após o login do cso
'Variável sPasso2 para a entrada da tela da referida operação
'Variável sValida para que seja validado se a tela da operação foi carregada
'Variável iValida recebe o numero da linha aonde o texto deve estar
'==============================================================================
        
On Error GoTo errHandler
    
    Dim sTela   As String
    Dim vValida As Variant
    Dim iIndex  As Integer
                        
    With Me
        
        'Primeiro preenchimento
        .IntoText 23, 29, sPasso1, 5, True
        .gravaLog "Inicio Operação 1", sPath
        
'        .IntoText 1, 1, "A", 5, True
'        .gravaLog "Inicio Operação 1-1", sPath
        
        'Segundo Preenchimento
        .IntoText 1, 1, sPasso2, 5, True
        .gravaLog "Inicio Operação 2", sPath
        
        'Busca a informação da Tela conforme parametro informado da Linha
        sTela = .RetTextoTela(iValida, 1, 80)
                
        vValida = Split(sValida, ";")
        For iIndex = 0 To UBound(vValida)
        
            'Valida se a tela está correta
            If InStr(UCase(sTela), UCase(vValida(iIndex))) > 0 Then
                acessoOperacao = True
                Exit Function
            Else
                acessoOperacao = False
            End If
        Next iIndex
        
        Exit Function
        
    End With

errHandler:
    
    cPw.gravaLog "Erro acessoOperacao!", sPath
    acessoOperacao = False
    'MsgBox Err.Number & vbCrLf & Err.Description
End Function



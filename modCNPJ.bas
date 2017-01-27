Attribute VB_Name = "modCNPJ"
Function VerificaCNPJ(CNPJ As String) As String

Dim intSoma, intSoma1, intSoma2, intInteiro As Long
Dim intNumero, intMais, i, intResto As Integer
Dim intDig1, intDig2 As Integer
Dim strcampo, strCaracter, StrConf, strCNPJ, strDigVer As String
Dim dblDivisao As Double
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0

strDigVer = Right(CNPJ, 2)
strcampo = Left(CNPJ, 8)
strCNPJ = Right(CNPJ, 6)
strCNPJ = Left(strCNPJ, 4)
strcampo = Right(strcampo, 4) & strCNPJ

For i = 2 To 9
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
'Separa os 4 primeiros d�gitos do CNPJ
strcampo = Left(CNPJ, 4)
For i = 2 To 5
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0

strcampo = Left(CNPJ, 8)
strCNPJ = Right(CNPJ, 6)
strCNPJ = Left(strCNPJ, 4)
strcampo = Right(strcampo, 3) & strCNPJ & intDig1

For i = 2 To 9
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
strcampo = Left(CNPJ, 5)
For i = 2 To 6
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
StrConf = intDig1 & intDig2
VerificaCNPJ = StrConf

If VerificaCNPJ = strDigVer Then
   MsgBox "CNPJ v�lido!", vbInformation

Else
   MsgBox "CNPJ inv�lido", vbCritical
End If

End Function

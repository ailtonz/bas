Attribute VB_Name = "AZS_ARQUIVOS"
Option Explicit

Public Function Saida(strConteudo As String, strArquivo As String)
    Open CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1
End Function

Public Function TestaExistenciaArquivo(ByVal caminhoArquivo As String) As Boolean
On Error Resume Next
    
    TestaExistenciaArquivo = IIf(Dir$(caminhoArquivo) <> "", True, False)

End Function

Sub copiarBanco()
Dim fs As Object
Dim oldPath As String, newPath As String

oldPath = pathWorkSheetAddress 'Folder file is located in
newPath = pathWorkSheetAddress 'Folder to copy file to

Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile oldPath & "dbVendedor.mdb", newPath & Protocolo & "_" & Range(NomeUsuario) & ".mdb" 'This file was an .xls file

Set fs = Nothing

End Sub

Public Function Protocolo() As String
    Protocolo = Right(Year(Now()), 2) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & "-" & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
End Function

Public Function DivisorDeTexto(Texto As String, divisor As String, Indice As Integer) As String
Dim matriz As Variant
    
    matriz = Array()
    matriz = Split(Texto, divisor)
    DivisorDeTexto = CStr(matriz(Indice))

End Function

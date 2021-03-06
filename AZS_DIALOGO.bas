Attribute VB_Name = "AZS_DIALOGO"
Option Explicit

Public Function SelecionarBanco() As String
Dim fd As Office.FileDialog
Dim strArq As String
    
    On Error GoTo SelecionarBanco_err
    
    'Di�logo de selecionar arquivo - Office
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "BDs do Access", "*.MDB;*.MDE"
    fd.Title = "Localize a fonte de dados"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        strArq = fd.SelectedItems(1)
    End If
        
    If strArq <> "" Then SelecionarBanco = strArq

SelecionarBanco_Fim:
    Exit Function

SelecionarBanco_err:
    MsgBox Err.Description
    Resume SelecionarBanco_Fim

End Function

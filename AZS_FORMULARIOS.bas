Attribute VB_Name = "AZS_FORMULARIOS"
Option Explicit

Public Function CarregarComboBox(BaseDeDados As String, cbo As ComboBox, strCampo As String, strSQL As String) As Boolean: CarregarComboBox = True
On Error GoTo CarregarComboBox_err
Dim dbOrcamento As DAO.Database
Dim rstCarregarComboBox As DAO.Recordset
Dim retVal As Variant

retVal = Dir(BaseDeDados)

If retVal = "" Then

    CarregarComboBox = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstCarregarComboBox = dbOrcamento.OpenRecordset(strSQL)
    
    cbo.Clear
    
    While Not rstCarregarComboBox.EOF
        cbo.AddItem rstCarregarComboBox.Fields(strCampo)
        rstCarregarComboBox.MoveNext
    Wend
        
    rstCarregarComboBox.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarComboBox = Nothing
    
End If

CarregarComboBox_Fim:
  
    Exit Function
CarregarComboBox_err:
    CarregarComboBox = False
    MsgBox Err.Description
    Resume CarregarComboBox_Fim
End Function

Public Function CarregarListBox(BaseDeDados As String, frm As UserForm, NomeLista As String, strCampo As String, strSQL As String)
On Error GoTo CarregarListBox_err

Dim dbOrcamento         As DAO.Database
Dim rstCarregarListBox   As DAO.Recordset
Dim retVal              As Variant

Dim ctrl                As Control

retVal = Dir(BaseDeDados)

If retVal = "" Then

    CarregarListBox = False
    
Else
    
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstCarregarListBox = dbOrcamento.OpenRecordset(strSQL)
    
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "ListBox" Then
            If ctrl.Name = NomeLista Then
                ctrl.Clear
                While Not rstCarregarListBox.EOF
                    ctrl.AddItem rstCarregarListBox.Fields(strCampo)
                    rstCarregarListBox.MoveNext
                Wend
            End If
        End If
    Next
    
    rstCarregarListBox.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstCarregarListBox = Nothing
    
End If

CarregarListBox_Fim:
  
    Exit Function
CarregarListBox_err:
    CarregarListBox = False
    MsgBox Err.Description
    Resume CarregarListBox_Fim
End Function

Public Function DesbloqueioDeFuncoes(BaseDeDados As String, frm As UserForm, strSQL As String, strCampo As String)
On Error GoTo DesbloqueioDeFuncoes_err

Dim dbOrcamento         As DAO.Database
Dim rstDesbloqueioDeFuncoes   As DAO.Recordset
Dim retVal              As Variant
Dim ctrl                As Control

retVal = Dir(BaseDeDados)

If retVal = "" Then

    DesbloqueioDeFuncoes = False
    
Else
        
    Set dbOrcamento = DBEngine.OpenDatabase(BaseDeDados, False, False, "MS Access;PWD=" & SenhaBanco)
    Set rstDesbloqueioDeFuncoes = dbOrcamento.OpenRecordset(strSQL)
        
    While Not rstDesbloqueioDeFuncoes.EOF
        For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            If Right(ctrl.Name, Len(ctrl.Name) - 3) = rstDesbloqueioDeFuncoes.Fields(strCampo) Then
                ctrl.Enabled = True
            End If
            End If
        Next
        rstDesbloqueioDeFuncoes.MoveNext
    Wend
    
    rstDesbloqueioDeFuncoes.Close
    dbOrcamento.Close
    
    Set dbOrcamento = Nothing
    Set rstDesbloqueioDeFuncoes = Nothing
    
End If

DesbloqueioDeFuncoes_Fim:
  
    Exit Function
DesbloqueioDeFuncoes_err:
    DesbloqueioDeFuncoes = False
    MsgBox Err.Description
    Resume DesbloqueioDeFuncoes_Fim
End Function

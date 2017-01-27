VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Grands"
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

Public Sub add(ByVal grand As grand, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add grand, key, before, after
End Sub
Public Function count() As Integer
    count = pCol.count
End Function
Public Function remove(ByVal index As Integer)
    pCol.remove index
End Function
Public Function grand(ByVal index As Variant) As grand
    Set grand = pCol.item(index)
End Function


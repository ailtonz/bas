Attribute VB_Name = "Module1"
Sub try()
Dim arrayFields(100)
Dim i As Integer
Dim cCont As Control
i = 0
    For Each cCont In frmThresh.grpTxt1.Controls
        'If i = Main.numThresh Then
       ' Exit Sub
        'End If
        If TypeName(cCont) = "TextBox" Then
            cCont.Visible = True
        End If
     Next cCont


End Sub


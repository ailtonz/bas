Attribute VB_Name = "Module2"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Written by: Gleb Gudzenko
' Date: May 25, 2007
' Description: Makes a list of sheets and displays the ones selected
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' M&T BANK CORPORATION
' PROGRAM PROPRIETARY POLICY
 'This software program and/or related material has been developed exclusively for use by M&T BANK CORPORATION and
 'its subsidiaries. Any unauthorized use, transfer or transmission of this software program or related materials is expressly
 'forbidden. This material is not to be duplicated or removed from these premises without the express written permission of
 'M&T Bank Corporation's Management.

Option Explicit

Sub HideSheets()
    Dim tmpInt As Integer
    On Error Resume Next
    For tmpInt = 2 To Sheets.count
    Worksheets(tmpInt).Visible = xlSheetVeryHidden
    Next
End Sub
 
Sub ShowSheets()
Dim tmpInt As Integer
Application.ScreenUpdating = True

With MTBIA.ListBox1
    For tmpInt = 1 To .ListCount
        On Error Resume Next
        If .Selected(tmpInt) = True Then
            Worksheets(MTBIA.ListBox1.List(tmpInt)).Visible = xlSheetVisible
        End If
    Next
End With
End Sub
 
Sub CREATELIST()
Dim count, count2 As Integer
count2 = 1
For count = 2 To Sheets.count
    ThisWorkbook.Worksheets(1).Cells(count2, 1).Value = Sheets(count).Name
    count2 = count2 + 1
Next count
End Sub


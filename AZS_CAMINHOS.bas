Attribute VB_Name = "AZS_CAMINHOS"
Option Explicit

Public Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Public Function pathWorkSheetAddress() As String
    pathWorkSheetAddress = ActiveWorkbook.path & "\"
End Function

Public Function pathWorkbookFullName() As String
    pathWorkbookFullName = ActiveWorkbook.FullName
End Function



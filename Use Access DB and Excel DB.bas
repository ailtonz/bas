Attribute VB_Name = "Module1"


Public Sub GetData()
    Dim dbs As DAO.Database
    Dim rstEmployees As Recordset
    Dim fldEmployees As Field
    Dim intCount As Integer
    Dim empLname As String
    Dim empFname As String
    Dim empTitle As String

    intCount = 1
    
    Set dbs = DBEngine(0).OpenDatabase("C:\Program Files" _
        & "\Microsoft Office\Office\Samples\Northwind.mdb")

    Set rstEmployees = dbs.OpenRecordset("Employees", dbOpenTable)

    ' Set header values for sheet 1.
    With Worksheets("Sheet1").Rows(9)
        .Font.Bold = True
        .Cells(1, 5).Value = "Last Name"
        .Cells(1, 6).Value = "First Name"
        .Cells(1, 7).Value = "Job Title"
    End With

    ' Loop through all records, sending selected fields to AddToSheet
    ' function one row at a time.
    Do Until rstEmployees.EOF
        Set fldEmployees = rstEmployees.Fields(1)  ' "LastName" field.
        empLname = fldEmployees.Value
        Set fldEmployees = rstEmployees.Fields(2)  ' "FirstName" field.
        empFname = fldEmployees.Value
        Set fldEmployees = rstEmployees.Fields(3)  ' "Title" field.
        empTitle = fldEmployees.Value
        intCount = intCount + 1
        Call AddToSheet(intCount, empLname, empFname, empTitle)
       
        rstEmployees.MoveNext
    Loop
    
    With Worksheets("Sheet1").Columns("E:G")
        .AutoFit

    End With
End Sub


Public Function AddToSheet(intCount As Integer, empLname As String, _
    empFname As String, empTitle As String)

' This function adds recordset data to the sheet one row at a time.
' Values are passed from GetData procedure.
' The intCount variable is incremented one each pass so a new
' row is used for each record.
    
    With Worksheets("Sheet1").Rows(intCount)
        .Cells(10, 5).Value = empLname
        .Cells(10, 6).Value = empFname
        .Cells(10, 7).Value = empTitle
    End With
    
End Function


Public Sub FastData()
    Dim dbs As DAO.Database
    Dim rstEmployees As Recordset
    Dim strSQL As String
    Dim intCount As Integer
    
    Set dbs = DBEngine(0).OpenDatabase("C:\Program Files" _
        & "\Microsoft Office\Office\Samples\Northwind.mdb")
    Set rstEmployees = dbs.OpenRecordset("SELECT Employees.LastName, " _
        & "Employees.FirstName, Employees.Title FROM Employees;")
    
    With rstEmployees
        Do While Not .EOF
            ActiveSheet.Cells(intCount + 21, 5) = !LastName
            ActiveSheet.Cells(intCount + 21, 6) = !FirstName
            ActiveSheet.Cells(intCount + 21, 7) = !Title
            intCount = intCount + 1
            .MoveNext
        Loop
    End With
End Sub



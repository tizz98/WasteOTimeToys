Imports System.IO

Module modMain
    Sub Main()
        Dim employeeContainer As New EmployeeContainer()
        Dim parser As Parser
        Dim dataFilePath As String = promptUser("Please enter the data file path: ")

        If Not isFileValid(dataFilePath) Then
            Console.WriteLine("The path: " & dataFilePath & " is invalid.")
            Console.ReadLine()
            End
        End If

        parser = New Parser(dataFilePath)
        employeeContainer.setEmployees(parser.parseFile())

        employeeContainer.quickPrint()

        Console.ReadLine()
    End Sub

    Function promptUser(promptMsg As String) As String
        Console.WriteLine(promptMsg)
        Return Console.ReadLine()
    End Function

    Function isFileValid(filePath As String)
        Try
            Dim f As FileStream = File.Open(filePath, FileMode.Open)
            f.Close()
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function
End Module

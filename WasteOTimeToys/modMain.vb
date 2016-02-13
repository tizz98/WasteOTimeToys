'------------------------------------------------------------
'-                  File Name: modMain.vb                   -
'-                 Part of Project: Assign5                 -
'------------------------------------------------------------
'-                Written By: Elijah Wilson                 -
'-                  Written On: 02/13/2016                  -
'------------------------------------------------------------
'- File Purpose:                                            -
'-                                                          -
'- The main file of this program that contains Sub Main     -
'- where the program starts executing from.                 -
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- This program takes a file as an input and then parses it -
'- for employee data. After gathering the data, it then     -
'- generates a report that is shown on the screen.          -
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- (None)                                                   -
'------------------------------------------------------------
Imports System.IO

Module modMain
    '------------------------------------------------------------
    '-                  Subprogram Name: Main                   -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This is the main subprogram and is where the program     -
    '- executes from.                                           -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dataFilePath - The file path of the input data file      -
    '- employeeContainer - An EmployeeContainer object that is  -
    '-                     used to hold all the employees       -
    '- parser - A Parser object that is used to parse the input -
    '-          file                                            -
    '- report - The Report object that is used to generate the  -
    '-          data report                                     -
    '------------------------------------------------------------
    Sub Main()
        Console.Title = "Waste O' Time Toys"
        Dim employeeContainer As New EmployeeContainer()
        Dim parser As Parser
        Dim report As Report
        Dim dataFilePath As String = promptUser("Please enter the data file path: ")

        If Not isFileValid(dataFilePath) Then
            Console.WriteLine("The path: " & dataFilePath & " is invalid.")
            Console.ReadLine()
            End
        End If

        parser = New Parser(dataFilePath)
        employeeContainer.setEmployees(parser.parseFile())
        report = New Report(employeeContainer)

        ' Show Report
        Console.Clear()
        Console.WriteLine(report.generateReport())

        Console.ReadLine()
    End Sub

    '------------------------------------------------------------
    '-                Function Name: promptUser                 -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- To prompt the user with a question and get their         -
    '- response.                                                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- promptMsg - The message to prompt the user with          -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A string containing what the user typed         -
    '------------------------------------------------------------
    Function promptUser(promptMsg As String) As String
        Console.WriteLine(promptMsg)
        Return Console.ReadLine()
    End Function

    '------------------------------------------------------------
    '-                Function Name: isFileValid                -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- This checks whether or not a supplied file path is valid -
    '- to read from.                                            -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- filePath - The file path to test                         -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- f - Temporary FileStream object, when trying to open the -
    '-     file                                                 -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String) - Whether or not the supplied file path is valid -
    '------------------------------------------------------------
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

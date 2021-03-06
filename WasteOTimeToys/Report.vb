﻿'------------------------------------------------------------
'-                   File Name: Report.vb                   -
'-                 Part of Project: Assign5                 -
'------------------------------------------------------------
'-                Written By: Elijah Wilson                 -
'-                  Written On: 02/13/2016                  -
'------------------------------------------------------------
'- File Purpose:                                            -
'-                                                          -
'- Contains the Report class which is used to generate      -
'- reports based off of employee data.                      -
'------------------------------------------------------------
Public Class Report
    Private Const NO_DATA_STR As String = "Sorry, the input file is empty, so no data is available to report."
    Private Const FULL_LINE_LENGTH As Integer = 80
    Private Const CORP_NAME As String = "Waste O' Time Toys"
    Private Const REPORT_NAME As String = "Sales Report By {0}"
    Private Const REPORT_TITLE_FMT_STR As String = "*** {0} ***"
    Private Const ORDER_REPORT_NAME As String = "Order"
    Private Const NAME_REPORT_NAME As String = "Name"

    Private Const NAME_WIDTH As Integer = 20
    Private Const ID_WIDTH As Integer = 5
    Private Const NUM_CATEGORIES As Integer = 5
    Private Const SALE_CATEGORY_WIDTH As Integer = (FULL_LINE_LENGTH - NAME_WIDTH - ID_WIDTH) / NUM_CATEGORIES
    Private SALE_CATEGORY_WIDTH_STR As String = CStr(SALE_CATEGORY_WIDTH)
    Private NAMED_HEADERS_FMT_STR As String = "{0,-" & CStr(NAME_WIDTH) & "}{1,-" & CStr(ID_WIDTH) &
        "}{2," & SALE_CATEGORY_WIDTH_STR & "}{3," &
        SALE_CATEGORY_WIDTH & "}{4," & SALE_CATEGORY_WIDTH & "}{5," &
        SALE_CATEGORY_WIDTH & "}{6," & SALE_CATEGORY_WIDTH & "}"
    Private EMP_LINE_FMT_STR As String = "{0,-" & CStr(NAME_WIDTH) & "}{1,-" & CStr(ID_WIDTH) &
        "}{2," & SALE_CATEGORY_WIDTH_STR & ":C2}{3," &
        SALE_CATEGORY_WIDTH & ":C2}{4," & SALE_CATEGORY_WIDTH & ":C2}{5," &
        SALE_CATEGORY_WIDTH & ":C2}{6," & SALE_CATEGORY_WIDTH & ":C2}"

    Private STAT_NAMED_HEADERS_FMT_STR As String = "{0,-10}{1,14}{2,14}{3,14}{4,14}"
    Private QTY_STAT_ROW_FMT_STR As String = "{0,-10}{1,14:N2}{2,14:N2}{3,14:N2}{4,14:N2}"
    Private SALES_STAT_ROW_FMT_STR As String = "{0,-10}{1,14:C2}{2,14:C2}{3,14:C2}{4,14:C2}"
    Private SALES_STAT_LAST_ROW_FMT_STR As String = "{0,-10}{1,14:C2}{2,14:C2}{3,14:C2}{4,14:C2}{5,14}"
    Private STAT_TOTAL_ROW_FMT_STR As String = "{0,-10}{1,14:C2}{2,14:C2}{3,14:C2}{4,14:C2}{5,14:C2}"

    Private employees As EmployeeContainer

    Private Enum SortType
        LastNameFirstName
        OrderId
    End Enum

    '------------------------------------------------------------
    '-                   Subprogram Name: New                   -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Creates a new Report object.                             -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- employees - An EmployeeContainer object to be set for    -
    '-             this Report object                           -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New(employees As EmployeeContainer)
        Me.employees = employees
    End Sub

    '------------------------------------------------------------
    '-              Function Name: generateReport               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Generates a report based off of the employee data.       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- retString - A String that is accumulated over time and   -
    '-             eventually returned                          -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - The report as a string based off of the         -
    '-          employee data                                   -
    '------------------------------------------------------------
    Public Function generateReport() As String
        Dim retString As String = ""

        If employees.Count() > 0 Then
            retString &= getReportHeader(ORDER_REPORT_NAME) & vbCrLf
            retString &= getNamedHeaders() & vbCrLf & getDividerLine() & vbCrLf
            retString &= getEmployeeLines(SortType.OrderId) & vbCrLf & vbCrLf

            retString &= getReportHeader(NAME_REPORT_NAME) & vbCrLf
            retString &= getNamedHeaders() & vbCrLf & getDividerLine() & vbCrLf
            retString &= getEmployeeLines(SortType.LastNameFirstName) & vbCrLf & vbCrLf

            retString &= getSalesSummary() & vbCrLf
            retString &= getAboveAverageEmployees() & vbCrLf
        Else
            retString &= centerString(NO_DATA_STR) & vbCrLf
        End If

        Return retString
    End Function

    '------------------------------------------------------------
    '-         Function Name: getAboveAverageEmployees          -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a string describing how many employees sold      -
    '- above the average sales level as well as a list of the   -
    '- employees first name and last name.                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- aboveAvgEmployees - The employees that had sales above   -
    '-                     average                              -
    '- empCount - The number of employees who sold above the    -
    '-            average sales amount                          -
    '- retString - A String that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A string representation of which employees sold -
    '-          above the average sales amount                  -
    '------------------------------------------------------------
    Private Function getAboveAverageEmployees() As String
        Dim aboveAvgEmployees As List(Of Employee) = employees.getAboveAverageEmployees()
        Dim empCount As Integer = aboveAvgEmployees.Count
        Dim retString As String = ""

        If empCount = 1 Then
            retString &= String.Format("There is 1 employee who sold above the average sales level.")
        Else
            retString &= String.Format("There are {0} employees who sold above the average sales level.",
                                        empCount)
        End If

        retString &= vbCrLf

        ' Only show this if there was at least 1 employee who sold above the average
        If empCount > 0 Then
            If empCount = 1 Then
                retString &= "The name of the above average selling employee is:"
            Else
                retString &= "The names of the above average selling employees in alphabetical order are:"
            End If

            retString &= vbCrLf & vbCrLf

            For Each employee As Employee In aboveAvgEmployees
                retString &= String.Format("{0} {1}", employee.firstName, employee.lastName) & vbCrLf
            Next
        End If

        Return retString
    End Function

    '------------------------------------------------------------
    '-              Function Name: getSalesSummary              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- A summary about the sales and quantity of products sold. -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- retString - A string that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A string representation of a summary of the     -
    '-          sales and quantity of each product sold         -
    '------------------------------------------------------------
    Private Function getSalesSummary() As String
        Dim retString As String = getDividerLine() & vbCrLf & centerString("Sales Statistics Summary")
        retString &= vbCrLf & getDividerLine() & vbCrLf & getQuantityStats() & vbCrLf
        retString &= getSalesStats() & vbCrLf

        Return retString
    End Function

    '------------------------------------------------------------
    '-             Function Name: getQuantityStats              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Shows the Low, Average & High quantities sold for each   -
    '- product.                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- retString - A string that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - Statistics about the quanity of each product    -
    '-          sold                                            -
    '------------------------------------------------------------
    Private Function getQuantityStats() As String
        Dim retString As String = centerString("Quantity Statistics") & vbCrLf
        retString &= String.Format(STAT_NAMED_HEADERS_FMT_STR, "", "Games", "Dolls", "Building", "Model") & vbCrLf

        retString &= String.Format(QTY_STAT_ROW_FMT_STR, "Low", employees.getGameQtyLow(), employees.getDollQtyLow(),
                                   employees.getBldgQtyLow(), employees.getMdlQtyLow()) & vbCrLf
        retString &= String.Format(QTY_STAT_ROW_FMT_STR, "Avg", employees.getGameQtyAvg(), employees.getDollQtyAvg(),
                                   employees.getBldgQtyAvg(), employees.getMdlQtyAvg()) & vbCrLf
        retString &= String.Format(QTY_STAT_ROW_FMT_STR, "High", employees.getGameQtyHigh(), employees.getDollQtyHigh(),
                                   employees.getBldgQtyHigh(), employees.getMdlQtyHigh()) & vbCrLf

        Return retString
    End Function

    '------------------------------------------------------------
    '-               Function Name: getSalesStats               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Shows the Low, Average & High amount of sales for each   -
    '- product.                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- retString - A string that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - Statistics about the sales of each product sold -
    '------------------------------------------------------------
    Private Function getSalesStats() As String
        Dim retString As String = centerString("Sales Statistics") & vbCrLf
        retString &= String.Format(STAT_NAMED_HEADERS_FMT_STR, "", "Games", "Dolls", "Building", "Model") & vbCrLf

        retString &= String.Format(SALES_STAT_ROW_FMT_STR, "Low", employees.getGameSalesLow(), employees.getDollSalesLow(),
                                   employees.getBuildingSalesLow(), employees.getModelSalesLow()) & vbCrLf
        retString &= String.Format(SALES_STAT_ROW_FMT_STR, "Avg", employees.getGameSalesAvg(), employees.getDollSalesAvg(),
                                   employees.getBuildingSalesAvg(), employees.getModelSalesAvg()) & vbCrLf
        retString &= String.Format(SALES_STAT_LAST_ROW_FMT_STR, "High", employees.getGameSalesHigh(), employees.getDollSalesHigh(),
                                   employees.getBuildingSalesHigh(), employees.getModelSalesHigh(), "** Total **") & vbCrLf

        retString &= getDividerLine() & vbCrLf
        retString &= String.Format(STAT_TOTAL_ROW_FMT_STR, "Total", employees.getGameSalesTotal(), employees.getDollSalesTotal(),
                                   employees.getBuildingSalesTotal(), employees.getModelSalesTotal(), employees.getTotalSales()) & vbCrLf

        Return retString
    End Function

    '------------------------------------------------------------
    '-             Function Name: getEmployeeLines              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Gets a string of all the employees' information, each on -
    '- a separate line.                                         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sort - Which type of sort to use to sort the employees   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- retString - A string that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - Information about each employee displayed on a  -
    '-          separate line                                   -
    '------------------------------------------------------------
    Private Function getEmployeeLines(Optional sort As SortType = SortType.LastNameFirstName) As String
        Dim retString As String = ""

        ' Make sure employees are sorted
        Select Case (sort)
            Case SortType.LastNameFirstName
                Me.employees.SortByLastNameFirstName()
            Case SortType.OrderId
                Me.employees.SortByOrderId()
        End Select

        For Each employee In Me.employees
            retString &= getEmployeeLine(employee) & vbCrLf
        Next

        Return retString
    End Function

    '------------------------------------------------------------
    '-              Function Name: getEmployeeLine              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Information about an Employee object as it will be       -
    '- displayed as a line item.                                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- employee - The Employee object to use for information    -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - Information about an employee to be displayed   -
    '-          as a line item                                  -
    '------------------------------------------------------------
    Private Function getEmployeeLine(employee As Employee) As String
        Return String.Format(EMP_LINE_FMT_STR, employee.fullName(), employee.orderId.ToString().PadLeft(3, "0"),
                             employee.gameSales, employee.dollSales, employee.buildingSales, employee.modelSales,
                             employee.totalSales)
    End Function

    '------------------------------------------------------------
    '-              Function Name: getReportHeader              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a header for a report, that is customized by the -
    '- input reportName                                         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- reportName - The name of the report that should be used  -
    '-              in the header                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- reportTitle - The title of the report as determined by a -
    '-               format string and the input reportName     -
    '- retString - A string that is accumulated and eventually  -
    '-             returned                                     -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - The report header                               -
    '------------------------------------------------------------
    Private Function getReportHeader(reportName As String) As String
        Dim reportTitle As String = String.Format(REPORT_TITLE_FMT_STR, String.Format(REPORT_NAME, reportName))
        Dim retString As String = centerString(CORP_NAME) & vbCrLf
        retString &= centerString(reportTitle) & vbCrLf
        retString &= centerString(getDividerLine(reportTitle.Length)) & vbCrLf

        Return retString
    End Function

    '------------------------------------------------------------
    '-              Function Name: getNamedHeaders              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a string with the names of the headers formatted -
    '- properly                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - The headers formatted correctly                 -
    '------------------------------------------------------------
    Private Function getNamedHeaders() As String
        Return String.Format(NAMED_HEADERS_FMT_STR, "Name", "Id", "Games", "Dolls", "Building", "Models", "Total")
    End Function

    '------------------------------------------------------------
    '-              Function Name: getDividerLine               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a dashed divider line of a certain number of     -
    '- characters, defaults to FULL_LINE_LENGTH                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- length - The number of characters for the divider        -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - The divider line                                -
    '------------------------------------------------------------
    Private Function getDividerLine(Optional length As Integer = FULL_LINE_LENGTH) As String
        Return StrDup(length, "-")
    End Function

    '------------------------------------------------------------
    '-               Function Name: centerString                -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Centers a string within a certain line length.           -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- stringToCenter - The string to be centered               -
    '- lineLength - How long the line is                        -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- String - A String centered with a certain line length    -
    '------------------------------------------------------------
    Private Function centerString(stringToCenter As String, Optional lineLength As Integer = FULL_LINE_LENGTH) As String
        Return String.Format("{0,-" & CStr(lineLength) & "}",
                             String.Format("{0," & (Math.Ceiling((lineLength + stringToCenter.Length) / 2)).ToString() & "}", stringToCenter))
    End Function
End Class

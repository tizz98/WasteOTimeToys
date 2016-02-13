Public Class Report
    Private Const FULL_LINE_LENGTH As Integer = 80
    Private Const CORP_NAME As String = "Waste O' Time Toys"
    Private Const REPORT_NAME As String = "Sales Report By Order"
    Private Const REPORT_TITLE_FMT_STR As String = "*** {0} ***"
    Private REPORT_TITLE As String = String.Format(REPORT_TITLE_FMT_STR, REPORT_NAME)

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

    Private STAT_NAMED_HEADERS_FMT_STR As String = "{0,-15}{1,10}{2,10}{3,10}{4,10}"
    Private QTY_STAT_ROW_FMT_STR As String = "{0,-15}{1,10:N2}{2,10:N2}{3,10:N2}{4,10:N2}"
    Private SALES_STAT_ROW_FMT_STR As String = "{0,-15}{1,10:C2}{2,10:C2}{3,10:C2}{4,10:C2}"
    Private SALES_STAT_LAST_ROW_FMT_STR As String = "{0,-15}{1,10:C2}{2,10:C2}{3,10:C2}{4,10:C2}{5,20}"
    Private STAT_TOTAL_ROW_FMT_STR As String = "{0,-15}{1,10:C2}{2,10:C2}{3,10:C2}{4,10:C2}{5,20:C2}"

    Private employees As EmployeeContainer

    Public Sub New(employees As EmployeeContainer)
        Me.employees = employees
    End Sub

    Public Function generateReport() As String
        Dim retString As String = getReportHeader() & vbCrLf
        retString &= getNamedHeaders() & vbCrLf & getDividerLine() & vbCrLf
        retString &= getEmployeeLines() & vbCrLf & getSalesSummary() & vbCrLf

        Return retString
    End Function

    Private Function getSalesSummary() As String
        Dim retString As String = getDividerLine() & vbCrLf & centerString("Sales Statistics Summary")
        retString &= vbCrLf & getDividerLine() & vbCrLf & getQuantityStats() & vbCrLf
        retString &= getSalesStats() & vbCrLf

        Return retString
    End Function

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

    Private Function getSalesStats() As String
        Dim retString As String = centerString("Sales Statistics") & vbCrLf
        retString &= String.Format(STAT_NAMED_HEADERS_FMT_STR, "", "Games", "Dolls", "Building", "Model") & vbCrLf

        retString &= String.Format(SALES_STAT_ROW_FMT_STR, "Low", 0, 0, 0, 0) & vbCrLf
        retString &= String.Format(SALES_STAT_ROW_FMT_STR, "Avg", 0, 0, 0, 0) & vbCrLf
        retString &= String.Format(SALES_STAT_ROW_FMT_STR, "High", 0, 0, 0, 0, centerString("**Total**", 20)) & vbCrLf

        retString &= getDividerLine() & vbCrLf
        retString &= String.Format(STAT_TOTAL_ROW_FMT_STR, "Total", 0, 0, 0, 0, 0) & vbCrLf

        Return retString
    End Function

    Private Function getEmployeeLines() As String
        Dim retString As String = ""

        ' Make sure employees are sorted
        Me.employees.Sort()

        For Each employee In Me.employees
            retString &= getEmployeeLine(employee) & vbCrLf
        Next

        Return retString
    End Function

    Private Function getEmployeeLine(employee As Employee) As String
        Return String.Format(EMP_LINE_FMT_STR, employee.fullName(), employee.orderId, employee.gameSales,
                             employee.dollSales, employee.buildingSales, employee.modelSales, employee.totalSales)
    End Function

    Private Function getReportHeader() As String
        Dim retString As String = centerString(CORP_NAME) & vbCrLf
        retString &= centerString(REPORT_TITLE) & vbCrLf
        retString &= centerString(getDividerLine(REPORT_TITLE.Length)) & vbCrLf

        Return retString
    End Function

    Private Function getNamedHeaders() As String
        Return String.Format(NAMED_HEADERS_FMT_STR, "Name", "Id", "Games", "Dolls", "Building", "Models", "Total")
    End Function

    Private Function getDividerLine(Optional length As Integer = FULL_LINE_LENGTH) As String
        Return StrDup(length, "-")
    End Function

    Private Function centerString(stringToCenter As String, Optional lineLength As Integer = FULL_LINE_LENGTH) As String
        Return String.Format("{0,-" & CStr(lineLength) & "}",
                             String.Format("{0," & (Math.Ceiling((lineLength + stringToCenter.Length) / 2)).ToString() & "}", stringToCenter))
    End Function
End Class

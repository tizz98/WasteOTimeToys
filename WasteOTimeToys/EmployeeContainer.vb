Public Class EmployeeContainer
    Implements IEnumerable

    Private employees As New List(Of Employee)

    Public Sub setEmployees(employees As List(Of Employee))
        Me.employees = employees
    End Sub

    Public Sub Add(newEmployee As Employee)
        employees.Add(newEmployee)
    End Sub

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return employees.GetEnumerator()
    End Function

    Public Sub quickPrint()
        For Each employee In Me
            Console.WriteLine(employee.ToString())
        Next
    End Sub

    Public Sub Sort()
        SortByLastNameFirstName()
    End Sub

    Public Sub SortByLastNameFirstName()
        Me.employees = (From emp In Me.employees
                        Order By emp.lastName, emp.firstName Ascending
                        Select emp).ToList()
    End Sub

    Public Sub SortByOrderId()
        Me.employees = (From emp In Me.employees
                        Order By emp.orderId Ascending
                        Select emp).ToList()
    End Sub

    Public Function Count() As Integer
        Return employees.Count
    End Function

    Public Function getGameQtyLow() As Single
        Return (From emp In employees Select emp.gameQuantity).Min()
    End Function

    Public Function getGameQtyHigh() As Single
        Return (From emp In employees Select emp.gameQuantity).Max()
    End Function

    Public Function getGameQtyAvg() As Single
        Return (From emp In employees Select emp.gameQuantity).Average()
    End Function

    Public Function getDollQtyLow() As Single
        Return (From emp In employees Select emp.dollQuantity).Min()
    End Function

    Public Function getDollQtyHigh() As Single
        Return (From emp In employees Select emp.dollQuantity).Max()
    End Function

    Public Function getDollQtyAvg() As Single
        Return (From emp In employees Select emp.dollQuantity).Average()
    End Function

    Public Function getBldgQtyLow() As Single
        Return (From emp In employees Select emp.buildingQuanity).Min()
    End Function

    Public Function getBldgQtyHigh() As Single
        Return (From emp In employees Select emp.buildingQuanity).Max()
    End Function

    Public Function getBldgQtyAvg() As Single
        Return (From emp In employees Select emp.buildingQuanity).Average()
    End Function

    Public Function getMdlQtyLow() As Single
        Return (From emp In employees Select emp.modelQuantity).Min()
    End Function

    Public Function getMdlQtyHigh() As Single
        Return (From emp In employees Select emp.modelQuantity).Max()
    End Function

    Public Function getMdlQtyAvg() As Single
        Return (From emp In employees Select emp.modelQuantity).Average()
    End Function

    Public Function getGameSalesLow() As Single
        Return (From emp In employees Select emp.gameSales).Min()
    End Function

    Public Function getGameSalesHigh() As Single
        Return (From emp In employees Select emp.gameSales).Max()
    End Function

    Public Function getGameSalesAvg() As Single
        Return (From emp In employees Select emp.gameSales).Average()
    End Function

    Public Function getDollSalesLow() As Single
        Return (From emp In employees Select emp.dollSales).Min()
    End Function

    Public Function getDollSalesHigh() As Single
        Return (From emp In employees Select emp.dollSales).Max()
    End Function

    Public Function getDollSalesAvg() As Single
        Return (From emp In employees Select emp.dollSales).Average()
    End Function

    Public Function getBuildingSalesHigh() As Single
        Return (From emp In employees Select emp.buildingSales).Max()
    End Function

    Public Function getBuildingSalesAvg() As Single
        Return (From emp In employees Select emp.buildingSales).Average()
    End Function

    Public Function getBuildingSalesLow() As Single
        Return (From emp In employees Select emp.buildingSales).Min()
    End Function

    Public Function getModelSalesHigh() As Single
        Return (From emp In employees Select emp.modelSales).Max()
    End Function

    Public Function getModelSalesAvg() As Single
        Return (From emp In employees Select emp.modelSales).Average()
    End Function

    Public Function getModelSalesLow() As Single
        Return (From emp In employees Select emp.modelSales).Min()
    End Function

    Public Function getBuildingSalesTotal() As Single
        Return (From emp In employees Select emp.buildingSales).Sum()
    End Function

    Public Function getGameSalesTotal() As Single
        Return (From emp In employees Select emp.gameSales).Sum()
    End Function

    Public Function getModelSalesTotal() As Single
        Return (From emp In employees Select emp.modelSales).Sum()
    End Function

    Public Function getDollSalesTotal() As Single
        Return (From emp In employees Select emp.dollSales).Sum()
    End Function

    Public Function getTotalSales() As Single
        Return getBuildingSalesTotal() + getGameSalesTotal() + getModelSalesTotal() + getDollSalesTotal()
    End Function

    Public Function getAboveAvgGameSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.gameSales > getGameSalesAvg()
                Select emp).ToList()
    End Function

    Public Function getAboveAvgDollSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.dollSales > getDollSalesAvg()
                Select emp).ToList()
    End Function

    Public Function getAboveAvgBldgSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.buildingSales > getBuildingSalesAvg()
                Select emp).ToList()
    End Function

    Public Function getAboveAvgModelSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.modelSales > getModelSalesAvg()
                Select emp).ToList()
    End Function

    Public Function getAboveAverageEmployees() As List(Of Employee)
        Dim tempEmployees As New List(Of Employee)

        tempEmployees.AddRange(getAboveAvgGameSalesEmployees())
        tempEmployees.AddRange(getAboveAvgDollSalesEmployees())
        tempEmployees.AddRange(getAboveAvgBldgSalesEmployees())
        tempEmployees.AddRange(getAboveAvgModelSalesEmployees())

        ' `Distinct` because an employee could be above avg in more than one product
        Return (From emp In tempEmployees
                Order By emp.lastName, emp.firstName
                Select emp
                Distinct).ToList()
    End Function
End Class

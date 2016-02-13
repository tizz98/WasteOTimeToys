Public Class EmployeeContainer
    Implements IEnumerable

    ' An 'Employee' is more like a single 'EmployeeOrder'...
    ' there can be multiple employees with the same id and name...
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
        Me.employees = (From emp In Me.employees
                        Order By emp.lastName, emp.firstName Ascending).ToList()
    End Sub

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
End Class

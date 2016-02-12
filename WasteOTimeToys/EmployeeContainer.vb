Public Class EmployeeContainer
    Implements IEnumerable

    Private employees As New List(Of Employee)

    Public Sub setEmployees(employees As List(Of Employee))
        Me.employees = employees
    End Sub

    Public Sub Add(newEmployee As Employee)
        If Not employeeInEmployees(newEmployee) Then
            employees.Add(newEmployee)
        Else
            Dim employeeInList As Employee = employees.Find(Function(e) e.id = newEmployee.id)

            mergeEmployees(employeeInList, newEmployee)
        End If
    End Sub

    Private Sub mergeEmployees(ByRef finalEmployee As Employee, employeeToMerge As Employee)
        ' Merge the sales and quantities from `employeeToMerge` with that of `finalEmployee`
        finalEmployee.gameSales += employeeToMerge.gameSales
        finalEmployee.gameQuantity += employeeToMerge.gameQuantity

        finalEmployee.dollSales += employeeToMerge.dollSales
        finalEmployee.dollQuantity += employeeToMerge.dollQuantity

        finalEmployee.buildingSales += employeeToMerge.buildingSales
        finalEmployee.buildingQuanity += employeeToMerge.buildingQuanity

        finalEmployee.modelSales += employeeToMerge.modelSales
        finalEmployee.modelQuantity += employeeToMerge.modelQuantity
    End Sub

    Private Function employeeInEmployees(employeeToCheck As Employee)
        If employees.Count > 0 Then
            Return employees.Find(Function(e) e.id = employeeToCheck.id) IsNot Nothing
        End If
        Return False
    End Function

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return employees.GetEnumerator()
    End Function

    Public Sub quickPrint()
        For Each employee In Me
            Console.WriteLine(employee.ToString())
        Next
    End Sub
End Class

Imports System.Reflection

Public Class Employee
    Public id As Integer
    Public firstName As String
    Public lastName As String
    Public orderId As Integer

    Public gameSales As Single
    Public gameQuantity As Integer

    Public dollSales As Single
    Public dollQuantity As Integer

    Public buildingSales As Single
    Public buildingQuanity As Integer

    Public modelSales As Single
    Public modelQuantity As Integer

    Public totalSales As Single

    Public Overrides Function toString() As String
        Dim fields As FieldInfo() = Me.GetType().GetFields()
        Dim accStr As String = ""
        Dim fmtStr As String = "<" & Me.GetType().FullName & "({0})>"

        For Each field As FieldInfo In fields
            If Not field.IsSpecialName Then
                accStr &= String.Format("{0}: {1}, ", field.Name, field.GetValue(Me))
            End If
        Next

        accStr = accStr.TrimEnd(" ").TrimEnd(",")

        Return String.Format(fmtStr, accStr)
    End Function

    Public Function fullName() As String
        Return String.Format("{0}, {1}", Me.lastName, Me.firstName)
    End Function
End Class

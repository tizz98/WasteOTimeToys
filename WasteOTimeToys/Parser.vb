Imports System.IO
Public Class Parser
    Private sReader As StreamReader

    Private Const FIRST_NAME_IDX As Integer = 0
    Private Const LAST_NAME_IDX As Integer = 1
    Private Const ORDER_ID_IDX As Integer = 2
    Private Const EMP_ID_IDX As Integer = 3
    Private Const GAME_AMT_IDX As Integer = 4
    Private Const GAME_QTY_IDX As Integer = 5
    Private Const DOLL_AMT_IDX As Integer = 6
    Private Const DOLL_QTY_IDX As Integer = 7
    Private Const BLDG_AMT_IDX As Integer = 8
    Private Const BLDG_QTY_IDX As Integer = 9
    Private Const MDL_AMT_IDX As Integer = 10
    Private Const MDL_QTY_IDX As Integer = 11

    Public Sub New(filePath As String)
        sReader = New StreamReader(filePath)
    End Sub

    Public Function parseFile() As List(Of Employee)
        Dim retList As New List(Of Employee)

        Do While sReader.Peek() >= 0
            retList.Add(getEmployeeFromLine(sReader.ReadLine()))
        Loop

        Return retList
    End Function

    Private Function getEmployeeFromLine(line As String) As Employee
        Dim employee As New Employee()
        Dim fields As String() = line.Split(" ")

        With employee
            .firstName = fields(FIRST_NAME_IDX)
            .lastName = fields(LAST_NAME_IDX)
            .orderId = fields(ORDER_ID_IDX)
            .id = fields(EMP_ID_IDX)
            .gameSales = fields(GAME_AMT_IDX)
            .gameQuantity = fields(GAME_QTY_IDX)
            .dollSales = fields(DOLL_AMT_IDX)
            .dollQuantity = fields(DOLL_QTY_IDX)
            .buildingSales = fields(BLDG_AMT_IDX)
            .buildingQuanity = fields(BLDG_QTY_IDX)
            .modelSales = fields(MDL_AMT_IDX)
            .modelQuantity = fields(MDL_QTY_IDX)
            .totalSales = .gameSales + .dollSales + .buildingSales + .modelSales
        End With

        Return employee
    End Function
End Class

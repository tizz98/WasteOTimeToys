'------------------------------------------------------------
'-               File Name: EmployeeContainer               -
'-                 Part of Project: Assign5                 -
'------------------------------------------------------------
'-                Written By: Elijah Wilson                 -
'-                  Written On: 02/13/2016                  -
'------------------------------------------------------------
'- File Purpose:                                            -
'-                                                          -
'- Contains the EmployeeContainer class. This class manages -
'- a List of Employee objects and provides several          -
'- functions to get aggregate data from this list.          -
'------------------------------------------------------------
Public Class EmployeeContainer
    Implements IEnumerable

    Private employees As New List(Of Employee)

    '------------------------------------------------------------
    '-              Subprogram Name: setEmployees               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Sets the input list of employees to the objects          -
    '- employees.                                               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- employees - A list of Employee objects to set for the    -
    '-             class                                        -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub setEmployees(employees As List(Of Employee))
        Me.employees = employees
    End Sub

    '------------------------------------------------------------
    '-               Function Name: GetEnumerator               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns an IEnumerator so the class can be looped over   -
    '- using a For Each loop.                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- IEnumerator - The employees IEnumerator object           -
    '------------------------------------------------------------
    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return employees.GetEnumerator()
    End Function

    '------------------------------------------------------------
    '-         Subprogram Name: SortByLastNameFirstName         -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Sorts the employees by their last name & first name in   -
    '- ascending order.                                         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub SortByLastNameFirstName()
        Me.employees = (From emp In Me.employees
                        Order By emp.lastName, emp.firstName Ascending
                        Select emp).ToList()
    End Sub

    '------------------------------------------------------------
    '-              Subprogram Name: SortByOrderId              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- Sorts the employees by their order id in ascending       -
    '- order.                                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub SortByOrderId()
        Me.employees = (From emp In Me.employees
                        Order By emp.orderId Ascending
                        Select emp).ToList()
    End Sub

    '------------------------------------------------------------
    '-                   Function Name: Count                   -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns employees.Count                                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Integer - How many employees are in the employees list   -
    '------------------------------------------------------------
    Public Function Count() As Integer
        Return employees.Count
    End Function

    '------------------------------------------------------------
    '-               Function Name: getGameQtyLow               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest game quantity                         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest game quantity                            -
    '------------------------------------------------------------
    Public Function getGameQtyLow() As Single
        Return (From emp In employees Select emp.gameQuantity).Min()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getGameQtyHigh               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest game quantity                        -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest game quantity                           -
    '------------------------------------------------------------
    Public Function getGameQtyHigh() As Single
        Return (From emp In employees Select emp.gameQuantity).Max()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getGameQtyAvg               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average game quantity                        -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average game quantity                           -
    '------------------------------------------------------------
    Public Function getGameQtyAvg() As Single
        Return (From emp In employees Select emp.gameQuantity).Average()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getDollQtyLow               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest doll quantity sold                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest doll quantity sold                       -
    '------------------------------------------------------------
    Public Function getDollQtyLow() As Single
        Return (From emp In employees Select emp.dollQuantity).Min()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getDollQtyHigh               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest doll quantity sold                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest doll quantity sold                      -
    '------------------------------------------------------------
    Public Function getDollQtyHigh() As Single
        Return (From emp In employees Select emp.dollQuantity).Max()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getDollQtyAvg               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average doll quantity sold                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average doll quantity sold                      -
    '------------------------------------------------------------
    Public Function getDollQtyAvg() As Single
        Return (From emp In employees Select emp.dollQuantity).Average()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getBldgQtyLow               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest building quantity sold                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest bulding quantity sold                    -
    '------------------------------------------------------------
    Public Function getBldgQtyLow() As Single
        Return (From emp In employees Select emp.buildingQuanity).Min()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getBldgQtyHigh               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest building quantity sold               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest bulding quantity sold                   -
    '------------------------------------------------------------
    Public Function getBldgQtyHigh() As Single
        Return (From emp In employees Select emp.buildingQuanity).Max()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getBldgQtyAvg               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average building quantity sold               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average bulding quantity sold                   -
    '------------------------------------------------------------
    Public Function getBldgQtyAvg() As Single
        Return (From emp In employees Select emp.buildingQuanity).Average()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getMdlQtyLow                -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest model quantity sold                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest model quantity sold                      -
    '------------------------------------------------------------
    Public Function getMdlQtyLow() As Single
        Return (From emp In employees Select emp.modelQuantity).Min()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getMdlQtyHigh               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest model quantity sold                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest model quantity sold                     -
    '------------------------------------------------------------
    Public Function getMdlQtyHigh() As Single
        Return (From emp In employees Select emp.modelQuantity).Max()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getMdlQtyAvg                -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average model quantity sold                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average model quantity sold                     -
    '------------------------------------------------------------
    Public Function getMdlQtyAvg() As Single
        Return (From emp In employees Select emp.modelQuantity).Average()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getGameSalesLow              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest game sales amount                     -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest game sales amount                        -
    '------------------------------------------------------------
    Public Function getGameSalesLow() As Single
        Return (From emp In employees Select emp.gameSales).Min()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getGameSalesHigh              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest game sales amount                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest game sales amount                       -
    '------------------------------------------------------------
    Public Function getGameSalesHigh() As Single
        Return (From emp In employees Select emp.gameSales).Max()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getGameSalesAvg              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average game sales amount                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average game sales amount                       -
    '------------------------------------------------------------
    Public Function getGameSalesAvg() As Single
        Return (From emp In employees Select emp.gameSales).Average()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getDollSalesLow              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest doll sales amount                     -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest doll sales amount                        -
    '------------------------------------------------------------
    Public Function getDollSalesLow() As Single
        Return (From emp In employees Select emp.dollSales).Min()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getDollSalesHigh              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest doll sales amount                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest doll sales amount                       -
    '------------------------------------------------------------
    Public Function getDollSalesHigh() As Single
        Return (From emp In employees Select emp.dollSales).Max()
    End Function

    '------------------------------------------------------------
    '-              Function Name: getDollSalesAvg              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average doll sales amount                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average doll sales amount                       -
    '------------------------------------------------------------
    Public Function getDollSalesAvg() As Single
        Return (From emp In employees Select emp.dollSales).Average()
    End Function

    '------------------------------------------------------------
    '-           Function Name: getBuildingSalesHigh            -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest building sales amount                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest building sales amount                   -
    '------------------------------------------------------------
    Public Function getBuildingSalesHigh() As Single
        Return (From emp In employees Select emp.buildingSales).Max()
    End Function

    '------------------------------------------------------------
    '-            Function Name: getBuildingSalesAvg            -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average building sales amount                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average building sales amount                   -
    '------------------------------------------------------------
    Public Function getBuildingSalesAvg() As Single
        Return (From emp In employees Select emp.buildingSales).Average()
    End Function

    '------------------------------------------------------------
    '-            Function Name: getBuildingSalesLow            -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest building sales amount                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest building sales amount                    -
    '------------------------------------------------------------
    Public Function getBuildingSalesLow() As Single
        Return (From emp In employees Select emp.buildingSales).Min()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getModelSalesHigh             -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the highest model sales amount                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Highest model sales amount                      -
    '------------------------------------------------------------
    Public Function getModelSalesHigh() As Single
        Return (From emp In employees Select emp.modelSales).Max()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getModelSalesAvg              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the average model sales amount                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Average model sales amount                      -
    '------------------------------------------------------------
    Public Function getModelSalesAvg() As Single
        Return (From emp In employees Select emp.modelSales).Average()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getModelSalesLow              -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the lowest model sales amount                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Lowest model sales amount                       -
    '------------------------------------------------------------
    Public Function getModelSalesLow() As Single
        Return (From emp In employees Select emp.modelSales).Min()
    End Function

    '------------------------------------------------------------
    '-           Function Name: getBuildingSalesTotal           -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the total amount of building sales               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Total amount of building sales                  -
    '------------------------------------------------------------
    Public Function getBuildingSalesTotal() As Single
        Return (From emp In employees Select emp.buildingSales).Sum()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getGameSalesTotal             -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the total amount of game sales                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Total amount of game sales                      -
    '------------------------------------------------------------
    Public Function getGameSalesTotal() As Single
        Return (From emp In employees Select emp.gameSales).Sum()
    End Function

    '------------------------------------------------------------
    '-            Function Name: getModelSalesTotal             -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the total amount of model sales                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Total amount of model sales                     -
    '------------------------------------------------------------
    Public Function getModelSalesTotal() As Single
        Return (From emp In employees Select emp.modelSales).Sum()
    End Function

    '------------------------------------------------------------
    '-             Function Name: getDollSalesTotal             -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the total amount of doll sales                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Total amount of doll sales                      -
    '------------------------------------------------------------
    Public Function getDollSalesTotal() As Single
        Return (From emp In employees Select emp.dollSales).Sum()
    End Function

    '------------------------------------------------------------
    '-               Function Name: getTotalSales               -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns the total amount of all sales                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Single - Total amount of all sales                       -
    '------------------------------------------------------------
    Public Function getTotalSales() As Single
        Return getBuildingSalesTotal() + getGameSalesTotal() + getModelSalesTotal() + getDollSalesTotal()
    End Function

    '------------------------------------------------------------
    '-       Function Name: getAboveAvgGameSalesEmployees       -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a list of employees that had game sales above    -
    '- the average                                              -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- List(Of Employee) - List of employees that had game      -
    '-                     sales above the average              -
    '------------------------------------------------------------
    Public Function getAboveAvgGameSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.gameSales > getGameSalesAvg()
                Select emp).ToList()
    End Function

    '------------------------------------------------------------
    '-       Function Name: getAboveAvgDollSalesEmployees       -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a list of employees that had doll sales above    -
    '- the average                                              -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- List(Of Employee) - List of employees that had doll      -
    '-                     sales above the average              -
    '------------------------------------------------------------
    Public Function getAboveAvgDollSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.dollSales > getDollSalesAvg()
                Select emp).ToList()
    End Function

    '------------------------------------------------------------
    '-       Function Name: getAboveAvgBldgSalesEmployees       -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a list of employees that had building sales      -
    '- above the average                                        -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- List(Of Employee) - List of employees that had building  -
    '-                     sales above the average              -
    '------------------------------------------------------------
    Public Function getAboveAvgBldgSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.buildingSales > getBuildingSalesAvg()
                Select emp).ToList()
    End Function

    '------------------------------------------------------------
    '-      Function Name: getAboveAvgModelSalesEmployees       -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a list of employees that had model sales above   -
    '- the average                                              -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- List(Of Employee) - List of employees that had model     -
    '-                     sales above the average              -
    '------------------------------------------------------------
    Public Function getAboveAvgModelSalesEmployees() As List(Of Employee)
        Return (From emp In employees
                Where emp.modelSales > getModelSalesAvg()
                Select emp).ToList()
    End Function

    '------------------------------------------------------------
    '-         Function Name: getAboveAverageEmployees          -
    '------------------------------------------------------------
    '-                Written By: Elijah Wilson                 -
    '-                  Written On: 02/13/2016                  -
    '------------------------------------------------------------
    '- Function Purpose:                                        -
    '-                                                          -
    '- Returns a list of employees that had sales above the     -
    '- average in any product                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- tempEmployees - A list of employees that is accumulated  -
    '-                 and used to get data from                -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- List(Of Employee) - List of employees that had sales     -
    '-                     above the average in any product     -
    '------------------------------------------------------------
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

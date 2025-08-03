Attribute VB_Name = "SalesOrderManagement"
Option Explicit

' Order structure
Public Type OrderRecord
    OrderID As Long
    CustomerID As Long
    CustomerName As String
    OrderDate As Date
    TotalAmount As Double
End Type

' Order detail structure
Public Type OrderDetailRecord
    DetailID As Long
    OrderID As Long
    ProductID As Long
    ProductName As String
    Quantity As Long
    UnitPrice As Double
    TotalPrice As Double
End Type

' Create new sales order
Public Function CreateSalesOrder(customerID As Long, orderDetails() As OrderDetailRecord, totalAmount As Double) As Long
    On Error GoTo ErrorHandler
    
    ' Start transaction
    gConn.BeginTrans
    
    ' Insert order header
    Dim orderSql As String
    orderSql = "INSERT INTO SalesOrders (CustomerID, OrderDate, TotalAmount) VALUES (" & _
               customerID & ", GETDATE(), " & totalAmount & "); SELECT SCOPE_IDENTITY() as OrderID"
    
    Dim rs As ADODB.Recordset
    Set rs = gConn.Execute(orderSql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        Dim orderID As Long
        orderID = rs.Fields("OrderID").Value
        rs.Close
        
        ' Insert order details
        Dim i As Long
        For i = LBound(orderDetails) To UBound(orderDetails)
            Dim detailSql As String
            detailSql = "INSERT INTO OrderDetails (OrderID, ProductID, Quantity, UnitPrice) VALUES (" & _
                       orderID & ", " & orderDetails(i).ProductID & ", " & _
                       orderDetails(i).Quantity & ", " & orderDetails(i).UnitPrice & ")"
            
            If Not ExecuteCommand(detailSql) Then
                gConn.RollbackTrans
                CreateSalesOrder = -1
                Exit Function
            End If
            
            ' Update inventory
            If Not RemoveInventory(orderDetails(i).ProductID, orderDetails(i).Quantity, "Main Warehouse") Then
                gConn.RollbackTrans
                CreateSalesOrder = -1
                Exit Function
            End If
        Next i
        
        ' Commit transaction
        gConn.CommitTrans
        CreateSalesOrder = orderID
    Else
        gConn.RollbackTrans
        CreateSalesOrder = -1
    End If
    Exit Function
    
ErrorHandler:
    If gConn.State = adStateOpen Then
        gConn.RollbackTrans
    End If
    MsgBox "Error creating sales order: " & Err.Description, vbCritical
    CreateSalesOrder = -1
End Function

' Get order by ID
Public Function GetOrder(orderID As Long) As OrderRecord
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT so.OrderID, so.CustomerID, c.CustomerName, so.OrderDate, so.TotalAmount " & _
          "FROM SalesOrders so " & _
          "INNER JOIN Customers c ON so.CustomerID = c.CustomerID " & _
          "WHERE so.OrderID = " & orderID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        GetOrder.OrderID = rs.Fields("OrderID").Value
        GetOrder.CustomerID = rs.Fields("CustomerID").Value
        GetOrder.CustomerName = rs.Fields("CustomerName").Value
        GetOrder.OrderDate = rs.Fields("OrderDate").Value
        GetOrder.TotalAmount = rs.Fields("TotalAmount").Value
        rs.Close
    Else
        GetOrder.OrderID = -1
        rs.Close
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting order: " & Err.Description, vbCritical
    GetOrder.OrderID = -1
End Function

' Get order details
Public Function GetOrderDetails(orderID As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT od.DetailID, od.OrderID, od.ProductID, p.ProductName, " & _
          "od.Quantity, od.UnitPrice, (od.Quantity * od.UnitPrice) as TotalPrice " & _
          "FROM OrderDetails od " & _
          "INNER JOIN Products p ON od.ProductID = p.ProductID " & _
          "WHERE od.OrderID = " & orderID & " " & _
          "ORDER BY od.DetailID"
    
    Set GetOrderDetails = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting order details: " & Err.Description, vbCritical
    Set GetOrderDetails = Nothing
End Function

' Get all orders
Public Function GetAllOrders() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT so.OrderID, so.CustomerID, c.CustomerName, so.OrderDate, so.TotalAmount " & _
          "FROM SalesOrders so " & _
          "INNER JOIN Customers c ON so.CustomerID = c.CustomerID " & _
          "ORDER BY so.OrderDate DESC"
    
    Set GetAllOrders = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting all orders: " & Err.Description, vbCritical
    Set GetAllOrders = Nothing
End Function

' Get orders by date range
Public Function GetOrdersByDateRange(startDate As Date, endDate As Date) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT so.OrderID, so.CustomerID, c.CustomerName, so.OrderDate, so.TotalAmount " & _
          "FROM SalesOrders so " & _
          "INNER JOIN Customers c ON so.CustomerID = c.CustomerID " & _
          "WHERE so.OrderDate BETWEEN '" & Format(startDate, "yyyy-mm-dd") & "' " & _
          "AND '" & Format(endDate, "yyyy-mm-dd") & "' " & _
          "ORDER BY so.OrderDate DESC"
    
    Set GetOrdersByDateRange = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting orders by date range: " & Err.Description, vbCritical
    Set GetOrdersByDateRange = Nothing
End Function

' Get orders by customer
Public Function GetOrdersByCustomer(customerID As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT so.OrderID, so.CustomerID, c.CustomerName, so.OrderDate, so.TotalAmount " & _
          "FROM SalesOrders so " & _
          "INNER JOIN Customers c ON so.CustomerID = c.CustomerID " & _
          "WHERE so.CustomerID = " & customerID & " " & _
          "ORDER BY so.OrderDate DESC"
    
    Set GetOrdersByCustomer = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting orders by customer: " & Err.Description, vbCritical
    Set GetOrdersByCustomer = Nothing
End Function

' Calculate order total
Public Function CalculateOrderTotal(orderDetails() As OrderDetailRecord) As Double
    On Error GoTo ErrorHandler
    
    Dim total As Double
    total = 0
    
    Dim i As Long
    For i = LBound(orderDetails) To UBound(orderDetails)
        total = total + (orderDetails(i).Quantity * orderDetails(i).UnitPrice)
    Next i
    
    CalculateOrderTotal = total
    Exit Function
    
ErrorHandler:
    MsgBox "Error calculating order total: " & Err.Description, vbCritical
    CalculateOrderTotal = 0
End Function

' Get sales summary
Public Function GetSalesSummary(startDate As Date, endDate As Date) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT " & _
          "COUNT(so.OrderID) as OrderCount, " & _
          "SUM(so.TotalAmount) as TotalSales, " & _
          "AVG(so.TotalAmount) as AverageOrderValue, " & _
          "COUNT(DISTINCT so.CustomerID) as UniqueCustomers " & _
          "FROM SalesOrders so " & _
          "WHERE so.OrderDate BETWEEN '" & Format(startDate, "yyyy-mm-dd") & "' " & _
          "AND '" & Format(endDate, "yyyy-mm-dd") & "'"
    
    Set GetSalesSummary = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting sales summary: " & Err.Description, vbCritical
    Set GetSalesSummary = Nothing
End Function

' Get top selling products
Public Function GetTopSellingProducts(limit As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT p.ProductID, p.ProductName, p.Category, " & _
          "SUM(od.Quantity) as TotalQuantity, " & _
          "SUM(od.Quantity * od.UnitPrice) as TotalRevenue " & _
          "FROM OrderDetails od " & _
          "INNER JOIN Products p ON od.ProductID = p.ProductID " & _
          "GROUP BY p.ProductID, p.ProductName, p.Category " & _
          "ORDER BY TotalQuantity DESC"
    
    If limit > 0 Then
        sql = sql & " OFFSET 0 ROWS FETCH NEXT " & limit & " ROWS ONLY"
    End If
    
    Set GetTopSellingProducts = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting top selling products: " & Err.Description, vbCritical
    Set GetTopSellingProducts = Nothing
End Function

' Load orders to worksheet
Public Sub LoadOrdersToWorksheet(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ws.Cells.Clear
    
    ' Set headers
    ws.Cells(1, 1).Value = "Order ID"
    ws.Cells(1, 2).Value = "Customer ID"
    ws.Cells(1, 3).Value = "Customer Name"
    ws.Cells(1, 4).Value = "Order Date"
    ws.Cells(1, 5).Value = "Total Amount"
    
    ' Format headers
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(200, 200, 200)
    
    ' Get orders
    Dim rs As ADODB.Recordset
    Set rs = GetAllOrders()
    
    If Not rs Is Nothing Then
        Dim row As Long
        row = 2
        
        Do While Not rs.EOF
            ws.Cells(row, 1).Value = rs.Fields("OrderID").Value
            ws.Cells(row, 2).Value = rs.Fields("CustomerID").Value
            ws.Cells(row, 3).Value = rs.Fields("CustomerName").Value
            ws.Cells(row, 4).Value = rs.Fields("OrderDate").Value
            ws.Cells(row, 5).Value = rs.Fields("TotalAmount").Value
            row = row + 1
            rs.MoveNext
        Loop
        
        rs.Close
        
        ' Auto-fit columns
        ws.Columns("A:E").AutoFit
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading orders to worksheet: " & Err.Description, vbCritical
End Sub 
Attribute VB_Name = "CustomerManagement"
Option Explicit

' Customer structure
Public Type CustomerRecord
    CustomerID As Long
    CustomerName As String
    Email As String
    Phone As String
End Type

' Add new customer
Public Function AddCustomer(customerName As String, email As String, phone As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate email format
    If email <> "" And Not IsValidEmail(email) Then
        MsgBox "Invalid email format.", vbExclamation
        AddCustomer = False
        Exit Function
    End If
    
    Dim sql As String
    sql = "INSERT INTO Customers (CustomerName, Email, Phone) VALUES ('" & _
          Replace(customerName, "'", "''") & "', '" & _
          Replace(email, "'", "''") & "', '" & _
          Replace(phone, "'", "''") & "')"
    
    AddCustomer = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding customer: " & Err.Description, vbCritical
    AddCustomer = False
End Function

' Update existing customer
Public Function UpdateCustomer(customerID As Long, customerName As String, email As String, phone As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validate email format
    If email <> "" And Not IsValidEmail(email) Then
        MsgBox "Invalid email format.", vbExclamation
        UpdateCustomer = False
        Exit Function
    End If
    
    Dim sql As String
    sql = "UPDATE Customers SET CustomerName = '" & Replace(customerName, "'", "''") & "', " & _
          "Email = '" & Replace(email, "'", "''") & "', " & _
          "Phone = '" & Replace(phone, "'", "''") & "' " & _
          "WHERE CustomerID = " & customerID
    
    UpdateCustomer = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error updating customer: " & Err.Description, vbCritical
    UpdateCustomer = False
End Function

' Delete customer
Public Function DeleteCustomer(customerID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if customer has orders
    Dim checkSql As String
    checkSql = "SELECT COUNT(*) FROM SalesOrders WHERE CustomerID = " & customerID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(checkSql)
    
    If Not rs Is Nothing Then
        If rs.Fields(0).Value > 0 Then
            MsgBox "Cannot delete customer. They have existing orders.", vbExclamation
            rs.Close
            DeleteCustomer = False
            Exit Function
        End If
        rs.Close
    End If
    
    Dim sql As String
    sql = "DELETE FROM Customers WHERE CustomerID = " & customerID
    
    DeleteCustomer = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error deleting customer: " & Err.Description, vbCritical
    DeleteCustomer = False
End Function

' Get customer by ID
Public Function GetCustomer(customerID As Long) As CustomerRecord
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT CustomerID, CustomerName, Email, Phone FROM Customers WHERE CustomerID = " & customerID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        GetCustomer.CustomerID = rs.Fields("CustomerID").Value
        GetCustomer.CustomerName = rs.Fields("CustomerName").Value
        GetCustomer.Email = rs.Fields("Email").Value
        GetCustomer.Phone = rs.Fields("Phone").Value
        rs.Close
    Else
        GetCustomer.CustomerID = -1
        rs.Close
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting customer: " & Err.Description, vbCritical
    GetCustomer.CustomerID = -1
End Function

' Get all customers
Public Function GetAllCustomers() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT CustomerID, CustomerName, Email, Phone FROM Customers ORDER BY CustomerName"
    
    Set GetAllCustomers = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting customers: " & Err.Description, vbCritical
    Set GetAllCustomers = Nothing
End Function

' Search customers
Public Function SearchCustomers(searchTerm As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT CustomerID, CustomerName, Email, Phone FROM Customers " & _
          "WHERE CustomerName LIKE '%" & Replace(searchTerm, "'", "''") & "%' " & _
          "OR Email LIKE '%" & Replace(searchTerm, "'", "''") & "%' " & _
          "OR Phone LIKE '%" & Replace(searchTerm, "'", "''") & "%' " & _
          "ORDER BY CustomerName"
    
    Set SearchCustomers = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error searching customers: " & Err.Description, vbCritical
    Set SearchCustomers = Nothing
End Function

' Get customer order history
Public Function GetCustomerOrderHistory(customerID As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT so.OrderID, so.OrderDate, so.TotalAmount, " & _
          "COUNT(od.DetailID) as ItemCount " & _
          "FROM SalesOrders so " & _
          "LEFT JOIN OrderDetails od ON so.OrderID = od.OrderID " & _
          "WHERE so.CustomerID = " & customerID & " " & _
          "GROUP BY so.OrderID, so.OrderDate, so.TotalAmount " & _
          "ORDER BY so.OrderDate DESC"
    
    Set GetCustomerOrderHistory = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting customer order history: " & Err.Description, vbCritical
    Set GetCustomerOrderHistory = Nothing
End Function

' Get top customers by order value
Public Function GetTopCustomers(limit As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT c.CustomerID, c.CustomerName, c.Email, " & _
          "COUNT(so.OrderID) as OrderCount, " & _
          "SUM(so.TotalAmount) as TotalSpent " & _
          "FROM Customers c " & _
          "LEFT JOIN SalesOrders so ON c.CustomerID = so.CustomerID " & _
          "GROUP BY c.CustomerID, c.CustomerName, c.Email " & _
          "ORDER BY TotalSpent DESC"
    
    If limit > 0 Then
        sql = sql & " OFFSET 0 ROWS FETCH NEXT " & limit & " ROWS ONLY"
    End If
    
    Set GetTopCustomers = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting top customers: " & Err.Description, vbCritical
    Set GetTopCustomers = Nothing
End Function

' Validate email format
Private Function IsValidEmail(email As String) As Boolean
    ' Simple email validation
    Dim emailPattern As String
    emailPattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    ' For simplicity, we'll use a basic check
    IsValidEmail = InStr(email, "@") > 1 And InStr(email, ".") > InStr(email, "@")
End Function

' Load customers to worksheet
Public Sub LoadCustomersToWorksheet(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ws.Cells.Clear
    
    ' Set headers
    ws.Cells(1, 1).Value = "Customer ID"
    ws.Cells(1, 2).Value = "Customer Name"
    ws.Cells(1, 3).Value = "Email"
    ws.Cells(1, 4).Value = "Phone"
    
    ' Format headers
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Interior.Color = RGB(200, 200, 200)
    
    ' Get customers
    Dim rs As ADODB.Recordset
    Set rs = GetAllCustomers()
    
    If Not rs Is Nothing Then
        Dim row As Long
        row = 2
        
        Do While Not rs.EOF
            ws.Cells(row, 1).Value = rs.Fields("CustomerID").Value
            ws.Cells(row, 2).Value = rs.Fields("CustomerName").Value
            ws.Cells(row, 3).Value = rs.Fields("Email").Value
            ws.Cells(row, 4).Value = rs.Fields("Phone").Value
            row = row + 1
            rs.MoveNext
        Loop
        
        rs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading customers to worksheet: " & Err.Description, vbCritical
End Sub 
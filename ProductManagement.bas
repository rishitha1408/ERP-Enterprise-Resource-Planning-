Attribute VB_Name = "ProductManagement"
Option Explicit

' Product structure
Public Type ProductRecord
    ProductID As Long
    ProductName As String
    Category As String
    Price As Double
End Type

' Add new product
Public Function AddProduct(productName As String, category As String, price As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "INSERT INTO Products (ProductName, Category, Price) VALUES ('" & _
          Replace(productName, "'", "''") & "', '" & _
          Replace(category, "'", "''") & "', " & _
          price & ")"
    
    AddProduct = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding product: " & Err.Description, vbCritical
    AddProduct = False
End Function

' Update existing product
Public Function UpdateProduct(productID As Long, productName As String, category As String, price As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "UPDATE Products SET ProductName = '" & Replace(productName, "'", "''") & "', " & _
          "Category = '" & Replace(category, "'", "''") & "', " & _
          "Price = " & price & " " & _
          "WHERE ProductID = " & productID
    
    UpdateProduct = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error updating product: " & Err.Description, vbCritical
    UpdateProduct = False
End Function

' Delete product
Public Function DeleteProduct(productID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if product is used in orders
    Dim checkSql As String
    checkSql = "SELECT COUNT(*) FROM OrderDetails WHERE ProductID = " & productID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(checkSql)
    
    If Not rs Is Nothing Then
        If rs.Fields(0).Value > 0 Then
            MsgBox "Cannot delete product. It is used in existing orders.", vbExclamation
            rs.Close
            DeleteProduct = False
            Exit Function
        End If
        rs.Close
    End If
    
    ' Delete from inventory first
    Dim deleteInventorySql As String
    deleteInventorySql = "DELETE FROM Inventory WHERE ProductID = " & productID
    ExecuteCommand deleteInventorySql
    
    ' Delete product
    Dim sql As String
    sql = "DELETE FROM Products WHERE ProductID = " & productID
    
    DeleteProduct = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error deleting product: " & Err.Description, vbCritical
    DeleteProduct = False
End Function

' Get product by ID
Public Function GetProduct(productID As Long) As ProductRecord
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT ProductID, ProductName, Category, Price FROM Products WHERE ProductID = " & productID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        GetProduct.ProductID = rs.Fields("ProductID").Value
        GetProduct.ProductName = rs.Fields("ProductName").Value
        GetProduct.Category = rs.Fields("Category").Value
        GetProduct.Price = rs.Fields("Price").Value
        rs.Close
    Else
        GetProduct.ProductID = -1
        rs.Close
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting product: " & Err.Description, vbCritical
    GetProduct.ProductID = -1
End Function

' Get all products
Public Function GetAllProducts() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT ProductID, ProductName, Category, Price FROM Products ORDER BY ProductName"
    
    Set GetAllProducts = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting products: " & Err.Description, vbCritical
    Set GetAllProducts = Nothing
End Function

' Get products by category
Public Function GetProductsByCategory(category As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT ProductID, ProductName, Category, Price FROM Products " & _
          "WHERE Category = '" & Replace(category, "'", "''") & "' " & _
          "ORDER BY ProductName"
    
    Set GetProductsByCategory = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting products by category: " & Err.Description, vbCritical
    Set GetProductsByCategory = Nothing
End Function

' Search products
Public Function SearchProducts(searchTerm As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT ProductID, ProductName, Category, Price FROM Products " & _
          "WHERE ProductName LIKE '%" & Replace(searchTerm, "'", "''") & "%' " & _
          "OR Category LIKE '%" & Replace(searchTerm, "'", "''") & "%' " & _
          "ORDER BY ProductName"
    
    Set SearchProducts = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error searching products: " & Err.Description, vbCritical
    Set SearchProducts = Nothing
End Function

' Load products to worksheet
Public Sub LoadProductsToWorksheet(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ws.Cells.Clear
    
    ' Set headers
    ws.Cells(1, 1).Value = "Product ID"
    ws.Cells(1, 2).Value = "Product Name"
    ws.Cells(1, 3).Value = "Category"
    ws.Cells(1, 4).Value = "Price"
    
    ' Format headers
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Interior.Color = RGB(200, 200, 200)
    
    ' Get products
    Dim rs As ADODB.Recordset
    Set rs = GetAllProducts()
    
    If Not rs Is Nothing Then
        Dim row As Long
        row = 2
        
        Do While Not rs.EOF
            ws.Cells(row, 1).Value = rs.Fields("ProductID").Value
            ws.Cells(row, 2).Value = rs.Fields("ProductName").Value
            ws.Cells(row, 3).Value = rs.Fields("Category").Value
            ws.Cells(row, 4).Value = rs.Fields("Price").Value
            row = row + 1
            rs.MoveNext
        Loop
        
        rs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading products to worksheet: " & Err.Description, vbCritical
End Sub 
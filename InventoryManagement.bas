Attribute VB_Name = "InventoryManagement"
Option Explicit

' Inventory structure
Public Type InventoryRecord
    InventoryID As Long
    ProductID As Long
    ProductName As String
    Quantity As Long
    Location As String
End Type

' Add inventory record
Public Function AddInventoryRecord(productID As Long, quantity As Long, location As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check if product exists
    Dim checkSql As String
    checkSql = "SELECT COUNT(*) FROM Products WHERE ProductID = " & productID
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(checkSql)
    
    If Not rs Is Nothing Then
        If rs.Fields(0).Value = 0 Then
            MsgBox "Product does not exist.", vbExclamation
            rs.Close
            AddInventoryRecord = False
            Exit Function
        End If
        rs.Close
    End If
    
    ' Check if inventory record already exists for this product and location
    Dim existingSql As String
    existingSql = "SELECT InventoryID, Quantity FROM Inventory WHERE ProductID = " & productID & _
                  " AND Location = '" & Replace(location, "'", "''") & "'"
    
    Set rs = ExecuteQuery(existingSql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        ' Update existing record
        Dim updateSql As String
        updateSql = "UPDATE Inventory SET Quantity = Quantity + " & quantity & _
                    " WHERE InventoryID = " & rs.Fields("InventoryID").Value
        rs.Close
        AddInventoryRecord = ExecuteCommand(updateSql)
    Else
        ' Insert new record
        Dim insertSql As String
        insertSql = "INSERT INTO Inventory (ProductID, Quantity, Location) VALUES (" & _
                    productID & ", " & quantity & ", '" & Replace(location, "'", "''") & "')"
        rs.Close
        AddInventoryRecord = ExecuteCommand(insertSql)
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error adding inventory record: " & Err.Description, vbCritical
    AddInventoryRecord = False
End Function

' Update inventory quantity
Public Function UpdateInventoryQuantity(inventoryID As Long, newQuantity As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "UPDATE Inventory SET Quantity = " & newQuantity & " WHERE InventoryID = " & inventoryID
    
    UpdateInventoryQuantity = ExecuteCommand(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error updating inventory quantity: " & Err.Description, vbCritical
    UpdateInventoryQuantity = False
End Function

' Remove inventory (for sales)
Public Function RemoveInventory(productID As Long, quantity As Long, location As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check available quantity
    Dim checkSql As String
    checkSql = "SELECT InventoryID, Quantity FROM Inventory WHERE ProductID = " & productID & _
               " AND Location = '" & Replace(location, "'", "''") & "'"
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(checkSql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        Dim currentQuantity As Long
        currentQuantity = rs.Fields("Quantity").Value
        
        If currentQuantity < quantity Then
            MsgBox "Insufficient inventory. Available: " & currentQuantity & ", Requested: " & quantity, vbExclamation
            rs.Close
            RemoveInventory = False
            Exit Function
        End If
        
        Dim newQuantity As Long
        newQuantity = currentQuantity - quantity
        
        Dim updateSql As String
        updateSql = "UPDATE Inventory SET Quantity = " & newQuantity & _
                    " WHERE InventoryID = " & rs.Fields("InventoryID").Value
        
        rs.Close
        RemoveInventory = ExecuteCommand(updateSql)
    Else
        MsgBox "No inventory found for this product and location.", vbExclamation
        rs.Close
        RemoveInventory = False
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error removing inventory: " & Err.Description, vbCritical
    RemoveInventory = False
End Function

' Get inventory by product
Public Function GetInventoryByProduct(productID As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT i.InventoryID, i.ProductID, p.ProductName, i.Quantity, i.Location " & _
          "FROM Inventory i " & _
          "INNER JOIN Products p ON i.ProductID = p.ProductID " & _
          "WHERE i.ProductID = " & productID & " " & _
          "ORDER BY i.Location"
    
    Set GetInventoryByProduct = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting inventory by product: " & Err.Description, vbCritical
    Set GetInventoryByProduct = Nothing
End Function

' Get all inventory with product details
Public Function GetAllInventory() As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT i.InventoryID, i.ProductID, p.ProductName, p.Category, i.Quantity, i.Location " & _
          "FROM Inventory i " & _
          "INNER JOIN Products p ON i.ProductID = p.ProductID " & _
          "ORDER BY p.ProductName, i.Location"
    
    Set GetAllInventory = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting all inventory: " & Err.Description, vbCritical
    Set GetAllInventory = Nothing
End Function

' Get low stock items (quantity < threshold)
Public Function GetLowStockItems(threshold As Long) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT i.InventoryID, i.ProductID, p.ProductName, p.Category, i.Quantity, i.Location " & _
          "FROM Inventory i " & _
          "INNER JOIN Products p ON i.ProductID = p.ProductID " & _
          "WHERE i.Quantity < " & threshold & " " & _
          "ORDER BY i.Quantity, p.ProductName"
    
    Set GetLowStockItems = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting low stock items: " & Err.Description, vbCritical
    Set GetLowStockItems = Nothing
End Function

' Get inventory by location
Public Function GetInventoryByLocation(location As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT i.InventoryID, i.ProductID, p.ProductName, p.Category, i.Quantity, i.Location " & _
          "FROM Inventory i " & _
          "INNER JOIN Products p ON i.ProductID = p.ProductID " & _
          "WHERE i.Location = '" & Replace(location, "'", "''") & "' " & _
          "ORDER BY p.ProductName"
    
    Set GetInventoryByLocation = ExecuteQuery(sql)
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting inventory by location: " & Err.Description, vbCritical
    Set GetInventoryByLocation = Nothing
End Function

' Get total inventory value
Public Function GetTotalInventoryValue() As Double
    On Error GoTo ErrorHandler
    
    Dim sql As String
    sql = "SELECT SUM(i.Quantity * p.Price) as TotalValue " & _
          "FROM Inventory i " & _
          "INNER JOIN Products p ON i.ProductID = p.ProductID"
    
    Dim rs As ADODB.Recordset
    Set rs = ExecuteQuery(sql)
    
    If Not rs Is Nothing And Not rs.EOF Then
        If Not IsNull(rs.Fields("TotalValue").Value) Then
            GetTotalInventoryValue = rs.Fields("TotalValue").Value
        Else
            GetTotalInventoryValue = 0
        End If
        rs.Close
    Else
        GetTotalInventoryValue = 0
        rs.Close
    End If
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting total inventory value: " & Err.Description, vbCritical
    GetTotalInventoryValue = 0
End Function

' Load inventory to worksheet
Public Sub LoadInventoryToWorksheet(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear existing data
    ws.Cells.Clear
    
    ' Set headers
    ws.Cells(1, 1).Value = "Inventory ID"
    ws.Cells(1, 2).Value = "Product ID"
    ws.Cells(1, 3).Value = "Product Name"
    ws.Cells(1, 4).Value = "Category"
    ws.Cells(1, 5).Value = "Quantity"
    ws.Cells(1, 6).Value = "Location"
    
    ' Format headers
    ws.Range("A1:F1").Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(200, 200, 200)
    
    ' Get inventory
    Dim rs As ADODB.Recordset
    Set rs = GetAllInventory()
    
    If Not rs Is Nothing Then
        Dim row As Long
        row = 2
        
        Do While Not rs.EOF
            ws.Cells(row, 1).Value = rs.Fields("InventoryID").Value
            ws.Cells(row, 2).Value = rs.Fields("ProductID").Value
            ws.Cells(row, 3).Value = rs.Fields("ProductName").Value
            ws.Cells(row, 4).Value = rs.Fields("Category").Value
            ws.Cells(row, 5).Value = rs.Fields("Quantity").Value
            ws.Cells(row, 6).Value = rs.Fields("Location").Value
            row = row + 1
            rs.MoveNext
        Loop
        
        rs.Close
        
        ' Auto-fit columns
        ws.Columns("A:F").AutoFit
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading inventory to worksheet: " & Err.Description, vbCritical
End Sub 
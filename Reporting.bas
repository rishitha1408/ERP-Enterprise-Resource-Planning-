Attribute VB_Name = "Reporting"
Option Explicit

' Generate sales report
Public Sub GenerateSalesReport(startDate As Date, endDate As Date, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set title
    ws.Cells(1, 1).Value = "Sales Report"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ' Set date range
    ws.Cells(2, 1).Value = "Period: " & Format(startDate, "mm/dd/yyyy") & " to " & Format(endDate, "mm/dd/yyyy")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Get sales summary
    Dim summaryRs As ADODB.Recordset
    Set summaryRs = GetSalesSummary(startDate, endDate)
    
    If Not summaryRs Is Nothing And Not summaryRs.EOF Then
        ' Summary section
        ws.Cells(4, 1).Value = "Summary"
        ws.Cells(4, 1).Font.Bold = True
        ws.Cells(4, 1).Font.Size = 12
        
        ws.Cells(5, 1).Value = "Total Orders:"
        ws.Cells(5, 2).Value = summaryRs.Fields("OrderCount").Value
        
        ws.Cells(6, 1).Value = "Total Sales:"
        ws.Cells(6, 2).Value = Format(summaryRs.Fields("TotalSales").Value, "$#,##0.00")
        
        ws.Cells(7, 1).Value = "Average Order Value:"
        ws.Cells(7, 2).Value = Format(summaryRs.Fields("AverageOrderValue").Value, "$#,##0.00")
        
        ws.Cells(8, 1).Value = "Unique Customers:"
        ws.Cells(8, 2).Value = summaryRs.Fields("UniqueCustomers").Value
        
        summaryRs.Close
    End If
    
    ' Get detailed sales data
    Dim salesRs As ADODB.Recordset
    Set salesRs = GetOrdersByDateRange(startDate, endDate)
    
    If Not salesRs Is Nothing Then
        ' Detailed sales section
        ws.Cells(10, 1).Value = "Detailed Sales"
        ws.Cells(10, 1).Font.Bold = True
        ws.Cells(10, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(11, 1).Value = "Order ID"
        ws.Cells(11, 2).Value = "Customer Name"
        ws.Cells(11, 3).Value = "Order Date"
        ws.Cells(11, 4).Value = "Total Amount"
        
        ' Format headers
        ws.Range("A11:D11").Font.Bold = True
        ws.Range("A11:D11").Interior.Color = RGB(200, 200, 200)
        
        ' Data
        Dim row As Long
        row = 12
        
        Do While Not salesRs.EOF
            ws.Cells(row, 1).Value = salesRs.Fields("OrderID").Value
            ws.Cells(row, 2).Value = salesRs.Fields("CustomerName").Value
            ws.Cells(row, 3).Value = salesRs.Fields("OrderDate").Value
            ws.Cells(row, 4).Value = Format(salesRs.Fields("TotalAmount").Value, "$#,##0.00")
            row = row + 1
            salesRs.MoveNext
        Loop
        
        salesRs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating sales report: " & Err.Description, vbCritical
End Sub

' Generate inventory report
Public Sub GenerateInventoryReport(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set title
    ws.Cells(1, 1).Value = "Inventory Report"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Get total inventory value
    Dim totalValue As Double
    totalValue = GetTotalInventoryValue()
    
    ws.Cells(4, 1).Value = "Total Inventory Value:"
    ws.Cells(4, 2).Value = Format(totalValue, "$#,##0.00")
    ws.Cells(4, 2).Font.Bold = True
    ws.Cells(4, 2).Font.Size = 12
    
    ' Get low stock items
    Dim lowStockRs As ADODB.Recordset
    Set lowStockRs = GetLowStockItems(10) ' Threshold of 10
    
    If Not lowStockRs Is Nothing Then
        ws.Cells(6, 1).Value = "Low Stock Items (Quantity < 10)"
        ws.Cells(6, 1).Font.Bold = True
        ws.Cells(6, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(7, 1).Value = "Product Name"
        ws.Cells(7, 2).Value = "Category"
        ws.Cells(7, 3).Value = "Quantity"
        ws.Cells(7, 4).Value = "Location"
        
        ' Format headers
        ws.Range("A7:D7").Font.Bold = True
        ws.Range("A7:D7").Interior.Color = RGB(255, 200, 200) ' Light red for low stock
        
        ' Data
        Dim row As Long
        row = 8
        
        Do While Not lowStockRs.EOF
            ws.Cells(row, 1).Value = lowStockRs.Fields("ProductName").Value
            ws.Cells(row, 2).Value = lowStockRs.Fields("Category").Value
            ws.Cells(row, 3).Value = lowStockRs.Fields("Quantity").Value
            ws.Cells(row, 4).Value = lowStockRs.Fields("Location").Value
            row = row + 1
            lowStockRs.MoveNext
        Loop
        
        lowStockRs.Close
    End If
    
    ' Get all inventory
    Dim inventoryRs As ADODB.Recordset
    Set inventoryRs = GetAllInventory()
    
    If Not inventoryRs Is Nothing Then
        Dim startRow As Long
        startRow = row + 2
        
        ws.Cells(startRow, 1).Value = "Complete Inventory"
        ws.Cells(startRow, 1).Font.Bold = True
        ws.Cells(startRow, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(startRow + 1, 1).Value = "Product Name"
        ws.Cells(startRow + 1, 2).Value = "Category"
        ws.Cells(startRow + 1, 3).Value = "Quantity"
        ws.Cells(startRow + 1, 4).Value = "Location"
        
        ' Format headers
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Font.Bold = True
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Interior.Color = RGB(200, 200, 200)
        
        ' Data
        row = startRow + 2
        
        Do While Not inventoryRs.EOF
            ws.Cells(row, 1).Value = inventoryRs.Fields("ProductName").Value
            ws.Cells(row, 2).Value = inventoryRs.Fields("Category").Value
            ws.Cells(row, 3).Value = inventoryRs.Fields("Quantity").Value
            ws.Cells(row, 4).Value = inventoryRs.Fields("Location").Value
            row = row + 1
            inventoryRs.MoveNext
        Loop
        
        inventoryRs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating inventory report: " & Err.Description, vbCritical
End Sub

' Generate customer report
Public Sub GenerateCustomerReport(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set title
    ws.Cells(1, 1).Value = "Customer Report"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Get top customers
    Dim topCustomersRs As ADODB.Recordset
    Set topCustomersRs = GetTopCustomers(10) ' Top 10 customers
    
    If Not topCustomersRs Is Nothing Then
        ws.Cells(4, 1).Value = "Top 10 Customers by Total Spent"
        ws.Cells(4, 1).Font.Bold = True
        ws.Cells(4, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(5, 1).Value = "Customer Name"
        ws.Cells(5, 2).Value = "Email"
        ws.Cells(5, 3).Value = "Order Count"
        ws.Cells(5, 4).Value = "Total Spent"
        
        ' Format headers
        ws.Range("A5:D5").Font.Bold = True
        ws.Range("A5:D5").Interior.Color = RGB(200, 200, 200)
        
        ' Data
        Dim row As Long
        row = 6
        
        Do While Not topCustomersRs.EOF
            ws.Cells(row, 1).Value = topCustomersRs.Fields("CustomerName").Value
            ws.Cells(row, 2).Value = topCustomersRs.Fields("Email").Value
            ws.Cells(row, 3).Value = topCustomersRs.Fields("OrderCount").Value
            ws.Cells(row, 4).Value = Format(topCustomersRs.Fields("TotalSpent").Value, "$#,##0.00")
            row = row + 1
            topCustomersRs.MoveNext
        Loop
        
        topCustomersRs.Close
    End If
    
    ' Get all customers
    Dim customersRs As ADODB.Recordset
    Set customersRs = GetAllCustomers()
    
    If Not customersRs Is Nothing Then
        Dim startRow As Long
        startRow = row + 2
        
        ws.Cells(startRow, 1).Value = "All Customers"
        ws.Cells(startRow, 1).Font.Bold = True
        ws.Cells(startRow, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(startRow + 1, 1).Value = "Customer ID"
        ws.Cells(startRow + 1, 2).Value = "Customer Name"
        ws.Cells(startRow + 1, 3).Value = "Email"
        ws.Cells(startRow + 1, 4).Value = "Phone"
        
        ' Format headers
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Font.Bold = True
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Interior.Color = RGB(200, 200, 200)
        
        ' Data
        row = startRow + 2
        
        Do While Not customersRs.EOF
            ws.Cells(row, 1).Value = customersRs.Fields("CustomerID").Value
            ws.Cells(row, 2).Value = customersRs.Fields("CustomerName").Value
            ws.Cells(row, 3).Value = customersRs.Fields("Email").Value
            ws.Cells(row, 4).Value = customersRs.Fields("Phone").Value
            row = row + 1
            customersRs.MoveNext
        Loop
        
        customersRs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating customer report: " & Err.Description, vbCritical
End Sub

' Generate product performance report
Public Sub GenerateProductPerformanceReport(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set title
    ws.Cells(1, 1).Value = "Product Performance Report"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Get top selling products
    Dim topProductsRs As ADODB.Recordset
    Set topProductsRs = GetTopSellingProducts(20) ' Top 20 products
    
    If Not topProductsRs Is Nothing Then
        ws.Cells(4, 1).Value = "Top 20 Selling Products"
        ws.Cells(4, 1).Font.Bold = True
        ws.Cells(4, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(5, 1).Value = "Product Name"
        ws.Cells(5, 2).Value = "Category"
        ws.Cells(5, 3).Value = "Total Quantity Sold"
        ws.Cells(5, 4).Value = "Total Revenue"
        
        ' Format headers
        ws.Range("A5:D5").Font.Bold = True
        ws.Range("A5:D5").Interior.Color = RGB(200, 200, 200)
        
        ' Data
        Dim row As Long
        row = 6
        
        Do While Not topProductsRs.EOF
            ws.Cells(row, 1).Value = topProductsRs.Fields("ProductName").Value
            ws.Cells(row, 2).Value = topProductsRs.Fields("Category").Value
            ws.Cells(row, 3).Value = topProductsRs.Fields("TotalQuantity").Value
            ws.Cells(row, 4).Value = Format(topProductsRs.Fields("TotalRevenue").Value, "$#,##0.00")
            row = row + 1
            topProductsRs.MoveNext
        Loop
        
        topProductsRs.Close
    End If
    
    ' Get all products
    Dim productsRs As ADODB.Recordset
    Set productsRs = GetAllProducts()
    
    If Not productsRs Is Nothing Then
        Dim startRow As Long
        startRow = row + 2
        
        ws.Cells(startRow, 1).Value = "All Products"
        ws.Cells(startRow, 1).Font.Bold = True
        ws.Cells(startRow, 1).Font.Size = 12
        
        ' Headers
        ws.Cells(startRow + 1, 1).Value = "Product ID"
        ws.Cells(startRow + 1, 2).Value = "Product Name"
        ws.Cells(startRow + 1, 3).Value = "Category"
        ws.Cells(startRow + 1, 4).Value = "Price"
        
        ' Format headers
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Font.Bold = True
        ws.Range("A" & (startRow + 1) & ":D" & (startRow + 1)).Interior.Color = RGB(200, 200, 200)
        
        ' Data
        row = startRow + 2
        
        Do While Not productsRs.EOF
            ws.Cells(row, 1).Value = productsRs.Fields("ProductID").Value
            ws.Cells(row, 2).Value = productsRs.Fields("ProductName").Value
            ws.Cells(row, 3).Value = productsRs.Fields("Category").Value
            ws.Cells(row, 4).Value = Format(productsRs.Fields("Price").Value, "$#,##0.00")
            row = row + 1
            productsRs.MoveNext
        Loop
        
        productsRs.Close
        
        ' Auto-fit columns
        ws.Columns("A:D").AutoFit
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating product performance report: " & Err.Description, vbCritical
End Sub

' Export data to CSV
Public Sub ExportToCSV(ws As Worksheet, filePath As String)
    On Error GoTo ErrorHandler
    
    ' Save as CSV
    ws.SaveAs filePath, xlCSV
    
    MsgBox "Data exported to: " & filePath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting to CSV: " & Err.Description, vbCritical
End Sub

' Create dashboard summary
Public Sub CreateDashboardSummary(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Clear worksheet
    ws.Cells.Clear
    
    ' Set title
    ws.Cells(1, 1).Value = "ERP Dashboard Summary"
    ws.Cells(1, 1).Font.Size = 18
    ws.Cells(1, 1).Font.Bold = True
    
    ws.Cells(2, 1).Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm")
    ws.Cells(2, 1).Font.Italic = True
    
    ' Get key metrics
    Dim totalInventoryValue As Double
    totalInventoryValue = GetTotalInventoryValue()
    
    ' Get recent sales (last 30 days)
    Dim recentSalesRs As ADODB.Recordset
    Set recentSalesRs = GetOrdersByDateRange(Date - 30, Date)
    
    Dim recentSalesTotal As Double
    recentSalesTotal = 0
    Dim recentOrderCount As Long
    recentOrderCount = 0
    
    If Not recentSalesRs Is Nothing Then
        Do While Not recentSalesRs.EOF
            recentSalesTotal = recentSalesTotal + recentSalesRs.Fields("TotalAmount").Value
            recentOrderCount = recentOrderCount + 1
            recentSalesRs.MoveNext
        Loop
        recentSalesRs.Close
    End If
    
    ' Get low stock count
    Dim lowStockRs As ADODB.Recordset
    Set lowStockRs = GetLowStockItems(10)
    
    Dim lowStockCount As Long
    lowStockCount = 0
    
    If Not lowStockRs Is Nothing Then
        Do While Not lowStockRs.EOF
            lowStockCount = lowStockCount + 1
            lowStockRs.MoveNext
        Loop
        lowStockRs.Close
    End If
    
    ' Display metrics
    ws.Cells(4, 1).Value = "Key Metrics"
    ws.Cells(4, 1).Font.Bold = True
    ws.Cells(4, 1).Font.Size = 14
    
    ws.Cells(5, 1).Value = "Total Inventory Value:"
    ws.Cells(5, 2).Value = Format(totalInventoryValue, "$#,##0.00")
    ws.Cells(5, 2).Font.Bold = True
    
    ws.Cells(6, 1).Value = "Recent Sales (30 days):"
    ws.Cells(6, 2).Value = Format(recentSalesTotal, "$#,##0.00")
    ws.Cells(6, 2).Font.Bold = True
    
    ws.Cells(7, 1).Value = "Recent Orders (30 days):"
    ws.Cells(7, 2).Value = recentOrderCount
    ws.Cells(7, 2).Font.Bold = True
    
    ws.Cells(8, 1).Value = "Low Stock Items (< 10):"
    ws.Cells(8, 2).Value = lowStockCount
    ws.Cells(8, 2).Font.Bold = True
    
    ' Auto-fit columns
    ws.Columns("A:B").AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating dashboard summary: " & Err.Description, vbCritical
End Sub 
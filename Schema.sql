-- ERP System Database Schema
-- This file contains the complete database structure for the ERP system

USE ERP_System;

-- Drop existing tables if they exist (for clean setup)
IF OBJECT_ID('OrderDetails', 'U') IS NOT NULL DROP TABLE OrderDetails;
IF OBJECT_ID('SalesOrders', 'U') IS NOT NULL DROP TABLE SalesOrders;
IF OBJECT_ID('Inventory', 'U') IS NOT NULL DROP TABLE Inventory;
IF OBJECT_ID('Customers', 'U') IS NOT NULL DROP TABLE Customers;
IF OBJECT_ID('Products', 'U') IS NOT NULL DROP TABLE Products;

-- Create Products table
CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100) NOT NULL,
    Category NVARCHAR(50) NOT NULL,
    Price DECIMAL(10, 2) NOT NULL,
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
);

-- Create Inventory table
CREATE TABLE Inventory (
    InventoryID INT PRIMARY KEY IDENTITY(1,1),
    ProductID INT FOREIGN KEY REFERENCES Products(ProductID),
    Quantity INT NOT NULL DEFAULT 0,
    Location NVARCHAR(100) NOT NULL,
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
);

-- Create Customers table
CREATE TABLE Customers (
    CustomerID INT PRIMARY KEY IDENTITY(1,1),
    CustomerName NVARCHAR(100) NOT NULL,
    Email NVARCHAR(100),
    Phone NVARCHAR(20),
    Address NVARCHAR(200),
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
);

-- Create SalesOrders table
CREATE TABLE SalesOrders (
    OrderID INT PRIMARY KEY IDENTITY(1,1),
    CustomerID INT FOREIGN KEY REFERENCES Customers(CustomerID),
    OrderDate DATETIME DEFAULT GETDATE(),
    TotalAmount DECIMAL(10,2) NOT NULL,
    Status NVARCHAR(20) DEFAULT 'Pending',
    CreatedDate DATETIME DEFAULT GETDATE(),
    ModifiedDate DATETIME DEFAULT GETDATE()
);

-- Create OrderDetails table
CREATE TABLE OrderDetails (
    DetailID INT PRIMARY KEY IDENTITY(1,1),
    OrderID INT FOREIGN KEY REFERENCES SalesOrders(OrderID),
    ProductID INT FOREIGN KEY REFERENCES Products(ProductID),
    Quantity INT NOT NULL,
    UnitPrice DECIMAL(10,2) NOT NULL,
    CreatedDate DATETIME DEFAULT GETDATE()
);

-- Create indexes for better performance
CREATE INDEX IX_Products_Category ON Products(Category);
CREATE INDEX IX_Inventory_ProductID ON Inventory(ProductID);
CREATE INDEX IX_Inventory_Location ON Inventory(Location);
CREATE INDEX IX_Customers_Email ON Customers(Email);
CREATE INDEX IX_SalesOrders_CustomerID ON SalesOrders(CustomerID);
CREATE INDEX IX_SalesOrders_OrderDate ON SalesOrders(OrderDate);
CREATE INDEX IX_OrderDetails_OrderID ON OrderDetails(OrderID);
CREATE INDEX IX_OrderDetails_ProductID ON OrderDetails(ProductID);

-- Create triggers for audit trail
GO

CREATE TRIGGER TR_Products_Update ON Products
AFTER UPDATE AS
BEGIN
    UPDATE Products 
    SET ModifiedDate = GETDATE()
    FROM Products p
    INNER JOIN inserted i ON p.ProductID = i.ProductID;
END

GO

CREATE TRIGGER TR_Inventory_Update ON Inventory
AFTER UPDATE AS
BEGIN
    UPDATE Inventory 
    SET ModifiedDate = GETDATE()
    FROM Inventory inv
    INNER JOIN inserted i ON inv.InventoryID = i.InventoryID;
END

GO

CREATE TRIGGER TR_Customers_Update ON Customers
AFTER UPDATE AS
BEGIN
    UPDATE Customers 
    SET ModifiedDate = GETDATE()
    FROM Customers c
    INNER JOIN inserted i ON c.CustomerID = i.CustomerID;
END

GO

CREATE TRIGGER TR_SalesOrders_Update ON SalesOrders
AFTER UPDATE AS
BEGIN
    UPDATE SalesOrders 
    SET ModifiedDate = GETDATE()
    FROM SalesOrders so
    INNER JOIN inserted i ON so.OrderID = i.OrderID;
END

GO

-- Create view for inventory with product details
CREATE VIEW vw_InventoryWithProducts AS
SELECT 
    i.InventoryID,
    i.ProductID,
    p.ProductName,
    p.Category,
    p.Price,
    i.Quantity,
    i.Location,
    (i.Quantity * p.Price) as InventoryValue
FROM Inventory i
INNER JOIN Products p ON i.ProductID = p.ProductID;

-- Create view for order details with product and customer info
CREATE VIEW vw_OrderDetailsWithInfo AS
SELECT 
    od.DetailID,
    od.OrderID,
    od.ProductID,
    p.ProductName,
    p.Category,
    od.Quantity,
    od.UnitPrice,
    (od.Quantity * od.UnitPrice) as TotalPrice,
    so.CustomerID,
    c.CustomerName,
    so.OrderDate,
    so.TotalAmount as OrderTotal
FROM OrderDetails od
INNER JOIN Products p ON od.ProductID = p.ProductID
INNER JOIN SalesOrders so ON od.OrderID = so.OrderID
INNER JOIN Customers c ON so.CustomerID = c.CustomerID;

-- Create view for sales summary
CREATE VIEW vw_SalesSummary AS
SELECT 
    so.OrderID,
    so.CustomerID,
    c.CustomerName,
    so.OrderDate,
    so.TotalAmount,
    COUNT(od.DetailID) as ItemCount,
    YEAR(so.OrderDate) as OrderYear,
    MONTH(so.OrderDate) as OrderMonth
FROM SalesOrders so
INNER JOIN Customers c ON so.CustomerID = c.CustomerID
LEFT JOIN OrderDetails od ON so.OrderID = od.OrderID
GROUP BY so.OrderID, so.CustomerID, c.CustomerName, so.OrderDate, so.TotalAmount;

PRINT 'ERP System database schema created successfully!'; 
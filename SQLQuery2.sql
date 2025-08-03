-- Create the database
CREATE DATABASE ERP_System;
GO

-- Switch to the database
USE ERP_System;
GO

-- Create tables
CREATE TABLE Products (
    ProductID INT PRIMARY KEY IDENTITY(1,1),
    ProductName NVARCHAR(100),
    Category NVARCHAR(50),
    Price DECIMAL(10, 2)
);

CREATE TABLE Inventory (
    InventoryID INT PRIMARY KEY IDENTITY(1,1),
    ProductID INT FOREIGN KEY REFERENCES Products(ProductID),
    Quantity INT,
    Location NVARCHAR(100)
);

CREATE TABLE Customers (
    CustomerID INT PRIMARY KEY IDENTITY(1,1),
    CustomerName NVARCHAR(100),
    Email NVARCHAR(100),
    Phone NVARCHAR(20)
);

CREATE TABLE SalesOrders (
    OrderID INT PRIMARY KEY IDENTITY(1,1),
    CustomerID INT FOREIGN KEY REFERENCES Customers(CustomerID),
    OrderDate DATETIME DEFAULT GETDATE(),
    TotalAmount DECIMAL(10,2)
);

CREATE TABLE OrderDetails (
    DetailID INT PRIMARY KEY IDENTITY(1,1),
    OrderID INT FOREIGN KEY REFERENCES SalesOrders(OrderID),
    ProductID INT FOREIGN KEY REFERENCES Products(ProductID),
    Quantity INT,
    UnitPrice DECIMAL(10,2)
);

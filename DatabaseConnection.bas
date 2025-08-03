Attribute VB_Name = "DatabaseConnection"
Option Explicit

' Global connection object
Public gConn As ADODB.Connection
Public gConnString As String

' Connection string - Update with your server details
Private Const CONNECTION_STRING As String = "Provider=SQLOLEDB;Data Source=YOUR_SERVER_NAME;Initial Catalog=ERP_System;Integrated Security=SSPI;"

' Initialize database connection
Public Function InitializeConnection() As Boolean
    On Error GoTo ErrorHandler
    
    Set gConn = New ADODB.Connection
    gConnString = CONNECTION_STRING
    gConn.Open gConnString
    
    InitializeConnection = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error connecting to database: " & Err.Description, vbCritical
    InitializeConnection = False
End Function

' Close database connection
Public Sub CloseConnection()
    If Not gConn Is Nothing Then
        If gConn.State = adStateOpen Then
            gConn.Close
        End If
        Set gConn = Nothing
    End If
End Sub

' Execute SQL query and return recordset
Public Function ExecuteQuery(sqlQuery As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If gConn Is Nothing Or gConn.State <> adStateOpen Then
        If Not InitializeConnection() Then
            Set ExecuteQuery = Nothing
            Exit Function
        End If
    End If
    
    rs.Open sqlQuery, gConn, adOpenStatic, adLockReadOnly
    Set ExecuteQuery = rs
    Exit Function
    
ErrorHandler:
    MsgBox "Error executing query: " & Err.Description, vbCritical
    Set ExecuteQuery = Nothing
End Function

' Execute SQL command (INSERT, UPDATE, DELETE)
Public Function ExecuteCommand(sqlCommand As String) As Boolean
    On Error GoTo ErrorHandler
    
    If gConn Is Nothing Or gConn.State <> adStateOpen Then
        If Not InitializeConnection() Then
            ExecuteCommand = False
            Exit Function
        End If
    End If
    
    gConn.Execute sqlCommand
    ExecuteCommand = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error executing command: " & Err.Description, vbCritical
    ExecuteCommand = False
End Function

' Test database connection
Public Sub TestConnection()
    If InitializeConnection() Then
        MsgBox "Database connection successful!", vbInformation
        CloseConnection
    Else
        MsgBox "Database connection failed!", vbCritical
    End If
End Sub

' Get connection status
Public Function IsConnected() As Boolean
    IsConnected = (Not gConn Is Nothing) And (gConn.State = adStateOpen)
End Function 
<div align="center">

## ADO Class


</div>

### Description

A small but handy ADO class to use with Classic ASP and Access. The class support all basic database calls like select queries, insert, update and delete by using the ADO Recordset Open method.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Terje Hauger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/terje-hauger.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/terje-hauger-ado-class__4-9227/archive/master.zip)





### Source Code

```
<%
'==========================================================
' MODULE:  cADO.asp
' AUTHOR:  www.u229.no
' CREATED: June 2005
'==========================================================
' COMMENT: A small but handy ADO class for use with Classic ASP and Access.
'         Covers most common database operations.
'==========================================================
' TODO: Extend the class to support MSSQL?
'==========================================================
' ROUTINES:
' - Public Property Let DatabaseType(s)
' - Public Property Let PathToDatabase(s)
' - Public Property Let Password(s)
' - Public Property Let UserName(s)
' - Public Property Let LockType(i)
' - Public Property Let CursorLocation(i)
' - Public Property Let CursorType(i)
' - Public Property Get ErrorMessage()
' - Private Sub Class_Initialize()
' - Private Sub Class_Terminate()
' - Public Function ExecuteSQL(sSQL, iMode)
'==========================================================
'// ADO CONSTANTS:
'---- CursorTypeEnum Values ----
Const adOpenUnspecified = -1
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
'---- LockTypeEnum Values ----
Const adLockUnspecified = -1
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4
'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3
'---- ObjectStateEnum Values ----
Const adStateOpen = &H1
' Get String contstant
Const adClipString = 2
'==========================================================
Class cADO
'==========================================================
Private m_oConn             '// Connection object
Private m_oRs               '// Recordset object
Private m_sPathToDatabase     '// Path to database
Private m_sDatabaseType       '// Type of database: Access (MSSQL not implemented)
Private m_sUserName          '// Database user name
Private m_sPassword           '// Database password
Private m_iLockType           '// Recordset Lock Type
Private m_iCursorLocation       '// Recordset Cursor Location
Private m_iCursorType         '// Recordset Cursor Type
Private m_sErrorMessage       '// Return a human readable error message
'// PROPERTIES
Public Property Let DatabaseType(s)
  m_sDatabaseType = s
End Property
Public Property Let PathToDatabase(s)
  m_sPathToDatabase = s
End Property
Public Property Let Password(s)
  m_sPassword = s
End Property
Public Property Let UserName(s)
  m_sUserName = s
End Property
Public Property Let LockType(i)
  m_iLockType = i
End Property
Public Property Let CursorLocation(i)
  m_iCursorLocation = i
End Property
Public Property Let CursorType(i)
  m_iCursorType = i
End Property
Public Property Get ErrorMessage()
  ErrorMessage = m_sErrorMessage
End Property
'--------------------------------------------------------------------------------------------------------
' Comment: Initialize the ADO Objects
'--------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
  On Error Resume Next
'---------------------------- Set default properties
  m_iLockType = adLockPessimistic
  m_iCursorLocation = adUseServer
  m_iCursorType = adOpenForwardOnly
  m_sUserName = ""
  m_sPassword = ""
  m_sErrorMessage = ""
  If IsEmpty(m_oConn) Then Set m_oConn = Server.CreateObject("ADODB.Connection")
  If IsEmpty(m_oRs) Then Set m_oRs = Server.CreateObject("ADODB.Recordset")
End Sub
'--------------------------------------------------------------------------------------------------------
' Comment:
'--------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
  On Error Resume Next
  m_oRs.Close
  Set m_oRs = Nothing
  Set m_oConn = Nothing
  Err.Clear
End Sub
'--------------------------------------------------------------------------------------------------------
' Comment: Retrieve the data and return them in requested form.
'--------------------------------------------------------------------------------------------------------
Public Function ExecuteSQL(sSQL, iMode)
  On Error Resume Next
'---------------------------- Simple check of user input
  If Not (IsNumeric(m_iLockType) Or IsNumeric(m_iCursorLocation) Or IsNumeric(m_iCursorType)) Then
    m_sErrorMessage = "Invalid parameter"
    Exit Function
  End If
'---------------------------- Select correct connection string and open the recordset
  With m_oConn
    If .State = adStateOpen Then .Close
    Select Case LCase(m_sDatabaseType)
      Case "access"
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & m_sPathToDatabase & _
		  ";User Id=" & m_sUserName & _
		  ";Password=" & m_sPassword
        '      Case "mssql"  '// Not implemented
      Case Else
        m_sErrorMessage = "Invalid or missing parameter for database type": Exit Function
    End Select
    .Open
  End With
'---------------------------- Set properties for the recordset object
  With m_oRs
    If .State = adStateOpen Then .Close
    .CursorType = ADOCursorType
    .CursorLocation = ADOCursorLocation
    .LockType = ADOLockType
    .ActiveConnection = m_oConn
    .Source = sSQL
    .Open
'---------------------------- Return the requested type of data
    Select Case iMode
      Case 1
        '// Use the GetRows method. This return a 2 dimensional array.
        ExecuteSQL = m_oRs.GetRows
      Case 2
        '// Return recordset. Set this function as a pointer to the recordset.
        Set ExecuteSQL = m_oRs
      Case 3
        '// Use the GetString method: GetString(StringFormat, NumRows, ColumnDelimiter, RowDelimiter, NullExpr)
        ExecuteSQL = m_oRs.GetString(adClipString)
      Case 4
        '// Just return a boolean.
        ExecuteSQL = (Err.Number = 0)
      Case Else
    End Select
  End With
End Function
'============================================================
End Class
'============================================================
'============================================================
' EXAMPLE 1: GetRows - 2 Dimensional array
'============================================================
If IsEmpty(oADO) Then Set oADO = New cADO
With oADO
  .DatabaseType = "Access"
  .PathToDatabase = "E:\folder1\folder2\MyDatabase.mdb"
  '// Returns a 2 dimensional array
  arrRecords = .ExecuteSQL("SELECT * FROM MyTable", 1)
End With
Set oADO = Nothing
For i = LBound(arrRecords) To UBound(arrRecords, 2)
  Response.Write arrRecords(1, i) & "<br>"
  Response.Write arrRecords(2, i) & "<br>"
  '// etc
Next
'============================================================
' EXAMPLE 2: Recordset
'============================================================
iCounter = 0
If IsEmpty(oADO) Then Set oADO = New cADO
If IsEmpty(oRs) Then Set oRs = Server.CreateObject("ADODB.Recordset")
With oADO
  .DatabaseType = "Access"
  .PathToDatabase = "E:\folder1\folder2\MyDatabase.mdb"
  '// Return a recordset object
  Set oRs = .ExecuteSQL("SELECT * FROM MyTable", 2)
End With
If Not IsEmpty(oRs) Or oRs.EOF Then
  Do While Not oRs.EOF
    Response.Write oRs(0) & "<br>"
    Response.Write oRs(1) & "<br>"
    oRs.MoveNext
  Loop
End If
'// Clean up
oRs.Close
Set oRs = Nothing
Set oADO = Nothing
'============================================================
' EXAMPLE 3: GetString
'============================================================
If IsEmpty(oADO) Then Set oADO = New cADO
With oADO
  .DatabaseType = "Access"
  .PathToDatabase = "E:\folder1\folder2\MyDatabase.mdb"
  sGetString = .ExecuteSQL("SELECT * FROM MyTable", 3)
End With
Set oADO = Nothing
'// vbTab is the default delimiter
arrValues = Split(sGetString, vbTab)
For i = LBound(arrValues) To UBound(arrValues)
  Response.Write arrValues(i) & "<br>"
Next
'============================================================
' EXAMPLE 4: Boolean
'============================================================
If IsEmpty(oADO) Then Set oADO = New cADO
With oADO
	.DatabaseType = "Access"
	.PathToDatabase = "E:\folder1\folder2\MyDatabase.mdb"
	'// Great for Update and delete operations
	bSuccess = .ExecuteSQL("DELETE FROM MyTable WHERE id=43", 4)
End With
Set oADO = Nothing
%>
```


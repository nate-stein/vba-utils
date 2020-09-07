VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' CLASS:    DATA WRAPPER
'           Provides a shell for accessing and manipulating ADODB Recordsets and databases.
' PROCESS:  This class retrieves and/or manipulates data via the following steps:
'           (1) Initiate() defines the ADODB.Connection.
'           (2) DefineRecord() defines the Recordset we will work with.
'           (3) With an opened Recordset, we can retrieve / manipulate data.
' SUBS:     OpenConnection
' FUNCT:    Query

'''''''''''''''''''''''''''''''''''''''
' ADODB objects
'''''''''''''''''''''''''''''''''''''''
Private mConnection As Object    'ADODB.Connection
Private mRecordset As Object     'ADODB.Recordset

'''''''''''''''''''''''''''''''''''''''
' ADODB Late-Binding Constants
'''''''''''''''''''''''''''''''''''''''
' Lock Types
Private Const m_LOCKTYPE_READONLY = 1
Private Const m_LOCKTYPE_PESSIMISTIC = 2
Private Const m_LOCKTYPE_OPTIMISTIC As Integer = 3
Private Const m_LOCKTYPE_BATCHOPTIMISTIC As Integer = 4

' Cursor Types
Private Const m_CURSORTYPE_OPENFORWARDONLY As Integer = 0
Private Const m_CURSORTYPE_OPENKEYSET As Integer = 1
Private Const m_CURSORTYPE_OPENDYNAMIC As Integer = 2
Private Const m_CURSORTYPE_OPENSTATIC As Integer = 3

' Cursor Location
Private Const m_CURSORLOCATION_USECLIENT As Integer = 3

' Search Direction
Private Const m_SEARCHDIRECTION_FORWARD As Integer = 0
Private Const m_SEARCHDIRECTION_BACKWARD As Integer = 1

' Search Starting Location
Private Const m_SEARCHSTART_CURRENT As Integer = 0
Private Const m_SEARCHSTART_FIRST As Integer = 1
Private Const m_SEARCHSTART_LAST As Integer = 2

' ADODB Error constants
Private Const m_ERROR_NOMATCHES As Long = 3021

'''''''''''''''''''''''''''''''''''''''
' Properties
'''''''''''''''''''''''''''''''''''''''
Private mConnectionOpen As Boolean
Private mConnectionString As String
Private mRecordsetDefined As Boolean
Private mReturnNullInsteadOfErrorOnFailedQuery As Boolean
'*****************************************************************************************
'*****************************************************************************************

Public Sub OpenConnection( _
   Optional ByVal dbPath As String = "", _
   Optional ByVal dbUserID As String = "", _
   Optional ByVal dbPassword As String = "")
'*********************************************************
' Opens ADODB.Connection to database.
' Create connection string if one hasn't already been provided to object.
'*********************************************************

   On Error GoTo ERR_HANDLER
   
   If Len(ConnectionString) = 0 Then
      ConnectionString = createConnectionString(dbPath, dbUserID, dbPassword)
   End If
      
   Set mConnection = CreateObject("ADODB.Connection")
   mConnection.Open ConnectionString
   mConnectionOpen = True
   
   Exit Sub
   
ERR_HANDLER:
   mConnectionOpen = False

End Sub

Public Function Query( _
   ByVal fields As Variant, _
   Optional ByVal sql As String = "") As Variant
'*********************************************************
' Returns value(s) contained in field(s).
' fields:   Can be a single String or an array. If it is an array, then function will return array
'           of corresponding values (in same order as their corresponding field elements in fields
'           variable).
' sql:      If no sql query is provided, function assumes that the caller has already defined a
'           Recordset from which we'll retrieve values.
'*********************************************************
      
   On Error GoTo ERR_HANDLER
   
   '''''''''''''''''''''''''''''''''''''''
   ' Ensure a Recordset has been defined, which would happen if either (i) caller passed a sql
   ' query to execute or (ii) if no sql query provided, an existing Recordset was previously
   ' defined.
   '''''''''''''''''''''''''''''''''''''''
   If Len(sql) > 0 Then DefineRecord sql, readOnly:=True
   
   If Not RecordsetDefined Then
      raiseDataWrapperError "Query()", "Recordset unsuccessfully defined."
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Return single result if fields is not an array.
   '''''''''''''''''''''''''''''''''''''''
   If Not IsArray(fields) Then
      Query = mRecordset.fields(fields).value
      Exit Function
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Create array for multiple return values since fields is an array.
   '''''''''''''''''''''''''''''''''''''''
   Dim results() As Variant
   ReDim results(0 To UBound(fields))
   Dim q As Integer
   For q = 0 To UBound(fields)
      results(q) = mRecordset.fields(fields(q)).value
   Next q
   
   Query = results
   Exit Function
      
ERR_HANDLER:
   If ReturnNullInsteadOfErrorOnFailedQuery Then
      Query = Null
      Exit Function
   Else:
      If Err.number = ig_ERR_DATA_WRAPPER Then
         iError_RaiseExisting
      Else: raiseDataWrapperError "GetFieldValue()"
      End If
   End If

End Function

Public Sub DefineRecord(ByVal sql As String, ByVal readOnly As Boolean)
'*********************************************************
' Opens Recordset.
'*********************************************************

   On Error GoTo ERR_HANDLER
      
   Dim result As Variant
   Set result = createRecordset(sql, readOnly)
   
   If IsNull(result) Then
      RecordsetDefined = False
      Exit Sub
   End If
   
   Set mRecordset = result
   RecordsetDefined = True
   Exit Sub
   
ERR_HANDLER:
   RecordsetDefined = False
      
End Sub

Private Function createRecordset(ByVal sql As String, ByVal readOnly As Boolean) As Variant
'*********************************************************
' Returns ADODB.Recordset object defined by sql query.
' Returns Null if an error is encountered or resulting Recordset is empty.
'*********************************************************

   On Error GoTo ERR_HANDLER
   
   Dim rs As Object
   Set rs = CreateObject("ADODB.Recordset")
   rs.CursorLocation = m_CURSORLOCATION_USECLIENT
   
   '''''''''''''''''''''''''''''''''''''''
   ' Define Recordset cursor settings depending on whether this Recordset is being created for
   ' editing purposes or only read-only.
   '''''''''''''''''''''''''''''''''''''''
   If readOnly Then
      rs.Open sql, mConnection, m_CURSORTYPE_OPENSTATIC, m_LOCKTYPE_READONLY
   Else: rs.Open sql, mConnection, m_CURSORTYPE_OPENDYNAMIC, m_LOCKTYPE_OPTIMISTIC
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Ensure we didn't open an empty Recordset (i.e. current record is positioned after the last
   ' record in the Recordset).
   '''''''''''''''''''''''''''''''''''''''
   If rs.EOF Then
      Set createRecordset = Null
   Else: Set createRecordset = rs
   End If
   Exit Function
   
ERR_HANDLER:
   Set createRecordset = Null

End Function

Private Function createConnectionString( _
   ByVal dbPath As String, _
   Optional ByVal dbUserID As String = "Admin", _
   Optional ByVal dbPassword As String = """""") As String
'*********************************************************
' Returns appropriate connection string to open ADODB.Connection to database located at dbPath.
'*********************************************************
   
   On Error GoTo ERR_HANDLER
   
   '''''''''''''''''''''''''''''''''''''''
   ' These are all components of the connection string whose values depend on the database type.
   '''''''''''''''''''''''''''''''''''''''
   Dim providerLanguage As String
   Dim engineNumber As Integer
   Dim extendedProperties As Variant ' only needed by Excel data sources.
      
   Select Case iFolder_GetFileExtension(dbPath)
      Case "accdb"
         providerLanguage = "Microsoft.ACE.OLEDB.12.0"
         engineNumber = 6
         extendedProperties = Null
      Case "mdb"
         providerLanguage = "Microsoft.Jet.OLEDB.4.0"
         engineNumber = 5
         extendedProperties = Null
      Case "xls"
         providerLanguage = "Microsoft.Jet.OLEDB.4.0"
         engineNumber = 5
         extendedProperties = ";Extended Properties=""Excel 8.0;HDR=YES"""
      Case "xlsx"
         providerLanguage = "Microsoft.ACE.OLEDB.12.0"
         engineNumber = 6
         extendedProperties = ";Extended Properties=""Excel 12.0 Xml;HDR=YES"""
      Case "xlsb"
         providerLanguage = "Microsoft.ACE.OLEDB.12.0"
         engineNumber = 6
         extendedProperties = ";Extended Properties=""Excel 12.0;HDR=YES"""
      Case "xlsm"
         providerLanguage = "Microsoft.ACE.OLEDB.12.0"
         engineNumber = 6
         extendedProperties = ";Extended Properties=""Excel 12.0 Macro;HDR=YES"""
      Case Else
         Err.Raise ig_ERR_DATA_WRAPPER, "createConnectionString()", "Unrecognized file extension."
   End Select
         
   Dim connString As String: connString = _
      "Provider=" & providerLanguage & ";User ID=" & dbUserID & ";Data Source=" & dbPath & ";" & _
      "Jet OLEDB:Database Password=" & dbPassword & ";Jet OLEDB:Engine Type=" & engineNumber
   
   If Not IsNull(extendedProperties) Then connString = connString & extendedProperties
   
   createConnectionString = connString
   Exit Function
   
ERR_HANDLER:
   raiseDataWrapperError "createConnectionString()"

End Function

Public Function AddRecord(ByVal fields As Variant, ByVal values As Variant) As Boolean
'*********************************************************
' Returns True if a new record is successfully added to current Recordset.
' Returns False if there is no open Recordset or an error is encountered adding a record to it.
'*********************************************************
      
   If Not RecordsetDefined Then
      AddRecord = False
      Exit Function
   End If
      
   On Error GoTo RETURN_FALSE
   
   With mRecordset
      .AddNew fields, values
      .Update
   End With
   DoEvents
   
   AddRecord = True
   Exit Function

RETURN_FALSE:
   AddRecord = False
     
End Function

Public Sub CloseRecord()
'*********************************************************
' Closes Recordset and release memory from object.
'*********************************************************

   On Error Resume Next
   mRecordset.Close
   Set mRecordset = Nothing
   Err.Clear
      
End Sub

Private Sub closeConnection()
'*********************************************************
' Close Connection and release its resources
'*********************************************************

   On Error Resume Next
   mConnection.Close
   Set mConnection = Nothing
   Err.Clear
      
End Sub

Private Sub Class_Terminate()
'*********************************************************
' Releases memory from ADODB-related objects.
'*********************************************************

   On Error Resume Next
   If Not (mRecordset Is Nothing) Then
      Call CloseRecord
      Set mRecordset = Nothing
   End If
   If Not (mConnection Is Nothing) Then
      Call closeConnection
      Set mConnection = Nothing
   End If
   Err.Clear
    
End Sub

Public Property Let ConnectionString(ByVal connString As String)

   mConnectionString = connString

End Property

Public Property Get ConnectionString() As String

   ConnectionString = mConnectionString

End Property

Public Property Let ReturnNullInsteadOfErrorOnFailedQuery(ByVal returnNull As Boolean)
'*********************************************************
' If returnNull True, then failed queries will return Null instead of raising errors.
'*********************************************************

   mReturnNullInsteadOfErrorOnFailedQuery = returnNull

End Property

Public Property Get ReturnNullInsteadOfErrorOnFailedQuery() As Boolean

   ReturnNullInsteadOfErrorOnFailedQuery = mReturnNullInsteadOfErrorOnFailedQuery

End Property

Public Property Get ConnectionIsOpen() As Boolean

   ConnectionIsOpen = mConnectionOpen

End Property

Private Property Let RecordsetDefined(ByVal isDefined As Boolean)

   mRecordsetDefined = isDefined

End Property

Public Property Get RecordsetDefined() As Boolean

   RecordsetDefined = mRecordsetDefined

End Property

Private Sub raiseDataWrapperError( _
   ByVal method As String, Optional ByVal descriptionOverride As String = "")
'*********************************************************
' Raises Err object for existing error.
' descriptionOverride: If none is provided, then the description that will be included in raised
'           error is the description from VBA's Err.Description. This is done because if the user
'           chooses to raise an IDataWrapper error somewhere along the code (where an error wasn't
'           being raised by VBA itself) it should have a reason that can be provided.
'*********************************************************

   Dim description As String
   If Len(descriptionOverride) > 0 Then
      description = descriptionOverride
   Else: description = Err.description
   End If
   
   Err.Raise ig_ERR_DATA_WRAPPER, method, description

End Sub

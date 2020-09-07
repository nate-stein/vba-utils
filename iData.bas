Attribute VB_Name = "iData"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   DATA UTILS
' PURPOSE:  Supporting methods for the IDataWrapper class and other data analysis.
' FUNCS:    CreateWrapper
'           FormatDateForAccess
'           IsSimpleType
'           IsSimpleNumericType
'*****************************************************************************************
'*****************************************************************************************

Public Function iData_CreateWrapper( _
   Optional ByVal dbPath As String = "", _
   Optional ByVal dbUserID As String = "Admin", _
   Optional ByVal dbPassword As String = """""", _
   Optional ByVal connString As String = "") As IDataWrapper
'*********************************************************
' Returns IDataWrapper Object.
' Serves as a Factory method for the IDataWrapper class.
'*********************************************************

   Dim data As IDataWrapper
   Set data = New IDataWrapper
   If Len(connString) > 0 Then
      data.ConnectionString = connString
      data.OpenConnection
   Else:
      data.OpenConnection dbPath, dbUserID, dbPassword
   End If

   Set iData_CreateWrapper = data

End Function

Public Function iData_FormatDateForAccess(ByVal dt As Date) As String
'*********************************************************
' Returns dt formatted in way that it can be injected into a SQL query.
'*********************************************************

   iData_FormatDateForAccess = "#" & Year(dt) & "-" & Format(dt, "MM") & "-" & Format(dt, "DD") & "#"

End Function

Public Function iData_GetDateFromUser( _
   Optional ByVal msg As String = "Enter date in following format: M/D/YYYY") As Date
'*********************************************************
' Returns a Date variable based on user input.
'*********************************************************
   
   Dim result As Variant
   result = InputBox(msg, "Enter Date")
      
   If Not IsDate(result) Then
      MsgBox "Not a valid date. Please try again.", , "Error"
      iData_GetDateFromUser = iData_GetDateFromUser(msg)
   Else:
      iData_GetDateFromUser = result
   End If

End Function

Public Function iData_IsSimpleType(ByVal v As Variant) As Boolean
'*********************************************************
' Returns True if v is not one of the following data types:
'     Array
'     Object
'     DataObject
'     UserDefinedType
'*********************************************************
   
   ' Test if v is an array. We can't just use VarType(v) = vbArray because the VarType of an array
   ' is vbArray + VarType -> a type of array element. E.g., the VarType of an Array of Longs is
   ' 8195 = vbArray + vbLong.
   If IsArray(v) Then
      iData_IsSimpleType = False
      Exit Function
   End If

   ' We must also explicitly check whether v is an object, rather than relying on VarType(v) to
   ' equal vbObject. The reason is that if v is an object and that object has a default property,
   ' VarType returns the data type of the default property. E.g., if v is an Excel.Range object
   ' pointing to cell A1, and A1 contains 12345, VarType(v) would return vbDouble, the since Value
   ' is the default property of an Excel.Range object and the default numeric type of Value in
   ' Excel is Double. Thus, in order to prevent this type of behavior with default properties, we
   ' test IsObject(v) to see if v is an object.
   If IsObject(v) Then
      iData_IsSimpleType = False
      Exit Function
   End If

   ' Now test type given we know v is not an array or object.
   Select Case VarType(v)
      Case vbArray, vbDataObject, vbObject, vbUserDefinedType
         iData_IsSimpleType = False
      Case Else
         iData_IsSimpleType = True
   End Select

End Function

Public Function iData_IsSimpleNumericType(ByVal v As Variant) As Boolean

   If Not iData_IsSimpleType(v) Then
      iData_IsSimpleNumericType = False
      Exit Function
   End If
   
   Select Case VarType(v)
      Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
         iData_IsSimpleNumericType = True
      Case vbVariant
         If IsNumeric(v) Then
            iData_IsSimpleNumericType = True
         Else: iData_IsSimpleNumericType = False
         End If
      Case Else
         iData_IsSimpleNumericType = False
   End Select
   
End Function

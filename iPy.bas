Attribute VB_Name = "iPy"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   PYTHON UTILS
' PURPOSE:  Tools to automate many tasks we perform in Excel when writing Python code.
' SYNTAX:   arr -> refers to arrays.
'*****************************************************************************************
'*****************************************************************************************

Public Function iPy_List( _
   ByVal vals As Variant, Optional ByVal noForceApostr As Boolean = False) As String
'*********************************************************
' Returns Python list containing all the values in vals.
' vals:     Can be Range or array.
'*********************************************************
   
   On Error GoTo errHandler
   
   Dim arr As Variant
   arr = ensureIsArray(vals)
   arr = ensureElementsProperlyFormatted(arr, noForceApostr)
   iPy_List = convertArrayToPythonListExpression(arr)
   
   Exit Function
   
errHandler:
   Select Case Err.number
      Case ig_ERR_PY
         iPy_List = Err.source & ": " & Err.description
      Case Else:
         iPy_List = "Error encountered: " & Err.description
   End Select
End Function

Public Function iPy_SQList(ByVal vals As Variant) As String
'*********************************************************
' Returns Python list containing all the values in vals.
' vals:     Can be Range or array.
'*********************************************************
    On Error GoTo errHandler
    
    Dim arr As Variant
    arr = ensureIsArray(vals)
    arr = ensureElementsProperlySQLFormatted(arr)
    iPy_SQList = convertArrayToPythonListExpression(arr)
    
    Exit Function
errHandler:
    Select Case Err.number
        Case ig_ERR_PY
            iPy_SQList = Err.source & ": " & Err.description
        Case Else:
            iPy_SQList = "Error encountered: " & Err.description
    End Select
End Function

Public Function iPy_DataForDF(ByVal rng As Range) As Variant
'*********************************************************
' Returns String representation of data that would be needed to construct a DataFrame from a dict
' of equal-length lists.
' Assumes the first row of rng contains the column headers.
'*********************************************************
   
   On Error GoTo errHandler
   
   Dim props As IT_RangeProperties
   props = iTools_GetRangeProperties(rng)
   
   ' Create dict entries by looping through columns and creating an entry for each column.
   Dim dictEntries() As Variant
   Dim col As Integer
   For col = props.LeftmostColumn To props.RightmostColumn
      Dim header As String
      header = "'" & cells(props.firstRow, col).value & "'"
      
      Dim listRng As Range
      Set listRng = Range(cells(props.firstRow + 1, col), cells(props.lastRow, col))
      Dim listExpr As String
      listExpr = iPy_List(listRng)
      
      Dim entry As String
      entry = header & ":" & listExpr
      iArr_AddValue dictEntries, entry
   Next col
   
   iPy_DataForDF = "{" & Join(dictEntries, ", ") & "}"
   
   Exit Function
   
errHandler:
   Select Case Err.number
      Case ig_ERR_PY
         iPy_DataForDF = Err.source & ": " & Err.description
      Case Else:
         iPy_DataForDF = "Error encountered: " & Err.description
   End Select
   
End Function

Public Function iPy_Dict( _
   ByVal keys As Variant, _
   ByVal vals As Variant, _
   Optional ByVal noForceApostrKeys As Boolean = False, _
   Optional ByVal noForceApostrVals As Boolean = False, _
   Optional ByVal doubleQuoteIfString As Boolean = False) As String
'*********************************************************
' Returns Python dict consisting of keys and corresponding vals.
' keys/vals: Can be Range or Array.
'*********************************************************

    On Error GoTo errHandler
    
    ' Get keys and vals into their final form.
    Dim trueKeys As Variant, trueVals As Variant
    trueKeys = ensureIsArray(keys)
    trueKeys = ensureElementsProperlyFormatted(trueKeys, noForceApostrKeys, doubleQuoteIfString)
    trueVals = ensureIsArray(vals)
    trueVals = ensureElementsProperlyFormatted(trueVals, noForceApostrVals, doubleQuoteIfString)
    
    Dim keyValuePairs As Variant
    keyValuePairs = createKeyValPairArray(trueKeys, trueVals)
    
    iPy_Dict = "{" & Join(keyValuePairs, ", ") & "}"
    
    Exit Function
   
errHandler:
   Select Case Err.number
      Case ig_ERR_PY
         iPy_Dict = Err.source & ": " & Err.description
      Case Else:
         iPy_Dict = "Error encountered: " & Err.description
   End Select

End Function

Public Function iPy_Matrix(ByVal rng As Range, Optional ByVal noForceApostr As Boolean = False)
'*********************************************************
' Returns matrix of data in rng, with rows for each row in rng and values for each row according
' to columns in range.
'*********************************************************
   
   On Error GoTo errHandler
   
   Dim props As IT_RangeProperties
   props = iTools_GetRangeProperties(rng)
   
   Dim matrixRows As Variant
   ReDim matrixRows(0 To props.RowCount - 1)
   
   Dim row As Integer, i As Integer
   i = 0
   For row = props.firstRow To props.lastRow
      Dim colRng As Range
      Set colRng = Range(cells(row, props.LeftmostColumn), cells(row, props.RightmostColumn))
      Dim matrixRow As String
      matrixRow = iPy_List(colRng, noForceApostr)
      matrixRows(i) = matrixRow
      i = i + 1
   Next row
   
   iPy_Matrix = "[" & Join(matrixRows, ", ") & "]"
   
   Exit Function
   
errHandler:
   Select Case Err.number
      Case ig_ERR_PY
         iPy_Matrix = Err.source & ": " & Err.description
      Case Else:
         iPy_Matrix = "Error encountered: " & Err.description
   End Select

End Function

Private Function createKeyValPairArray(ByVal keys As Variant, ByVal vals As Variant) As Variant
'*********************************************************
' Returns array of key:val pairs.
' keys/vals:   Must be array.
' Example: createKeyValPairArray(['A','B','C'], [1,2,3]) = ['A':1,'B':2,'C':3].
'*********************************************************

   On Error GoTo RAISE_ERR

   Dim result As Variant
   Dim q As Integer
   For q = 0 To UBound(keys) Step 1
      Dim pair As String
      pair = keys(q) & ": " & vals(q)
      iArr_AddValue result, pair
   Next q
   
   createKeyValPairArray = result
   Exit Function
   
RAISE_ERR:
   Err.Raise ig_ERR_PY, "createKeyValPairArray", Err.description

End Function

Private Function ensureIsArray(ByVal arr As Variant) As Variant
'*********************************************************
' Returns Array. If arr is a Range, it will be converted to an array. If arr is already an array,
' then arr will be returned unchanged.
' Raises error if arr is neither a Range or Arrray.
'*********************************************************
    If TypeName(arr) = "Range" Then
        ensureIsArray = iArr_FromRange(arr)
    ElseIf IsArray(arr) Then
        ensureIsArray = arr
    Else:
        Err.Raise ig_ERR_PY, "ensureIsArray", "arr needs to be Range or Array."
    End If
End Function

Private Function ensureElementsProperlySQLFormatted(ByVal arr As Variant) As Variant
'*********************************************************
' Returns arr after formatting elements according to its best guess on what
' data types they represent.
'*********************************************************
    Dim result As Variant, i As Integer, newval As Variant
    For i = 0 To UBound(arr) Step 1
        newval = castElementToSQLType(arr(i))
        iArr_AddValue result, newval
    Next i
    ensureElementsProperlySQLFormatted = result
End Function

Private Function castElementToSQLType(ByVal val As Variant) As Variant
    If IsNumeric(val) Then
        castElementToSQLType = val
        Exit Function
    End If
    
    Select Case LCase(val)
        Case "true", "false"
            castElementToSQLType = val
        Case Else
            castElementToSQLType = """" & val & """"
    End Select
End Function

Private Function ensureElementsProperlyFormatted( _
   ByVal arr As Variant, _
   ByVal noForceApostr As Boolean, _
   Optional ByVal doubleQuoteIfString As Boolean = False) As Variant
'*********************************************************
' Returns arr after ensuring the elements in arr are properly formatted based on their type. For
' example, if arr consists of Strings, then the single quotes will be added to each element.
' noForceApostr:
'           If True, then no apostrophes will be added to vals EVEN if they are determined to be
'           strings. If False, then normal logic will be used to determine whether to add
'           apostrophes or not.
' doubleQuoteIfString:
'           If True, then if apostrophes are added, it will be quotation marks
'           (not single apostrophes).
'*********************************************************
   If Not IsNumeric(arr(0)) And Not noForceApostr Then
      ensureElementsProperlyFormatted = addQuotesToStringArray(arr, doubleQuoteIfString)
   Else: ensureElementsProperlyFormatted = arr
   End If
End Function

Private Function convertArrayToPythonListExpression(ByVal arr As Variant) As String
'*********************************************************
' Returns Python list expression based on values in arr.
' Assumes that the values in arr are already properly formatted.
'*********************************************************
   
   convertArrayToPythonListExpression = "[" & Join(arr, ", ") & "]"

End Function

Private Function addQuotesToStringArray( _
    ByVal arr As Variant, _
    Optional ByVal useDoubleQuotes As Boolean = False) As Variant
'*********************************************************
' Returns copy of arr where each element is wrapped within single quotes.
' E.g. addQuotesToStringArray([3, 5, 7]) = ['3', '5', '7'].
'*********************************************************
    Dim result As Variant
    ReDim result(0 To UBound(arr))
    
    Dim bookmark As String
    If useDoubleQuotes Then
        bookmark = """"
    Else:
        bookmark = "'"
    End If
    
    Dim q As Integer
    For q = 0 To UBound(arr) Step 1
        result(q) = bookmark & arr(q) & bookmark
    Next q
    
    addQuotesToStringArray = result
End Function

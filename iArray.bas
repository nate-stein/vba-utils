Attribute VB_Name = "iArray"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   ARRAY UTILS
' WARNING:  Most of these methods assume that any input array was dimensioned from 0 to n - 1,
'           where n is the number of elements it contains (as opposed to 1 to n).
' METHODS:  AddValue
'           Combine1DArrays
'           ContainsValue
'           ConvertFromRange
'           CountUnique
'           DimensionNumber
'           Display
'           IsAllocated
'           MinimumValue
'           RemoveDuplicates
'           RemoveEmptyAndNullElementsFromEnd
'           ReverseInPlace
'           SortInPlace
'*****************************************************************************************
'*****************************************************************************************

Public Sub iArr_AddValue(ByRef arr As Variant, ByVal values As Variant)
'*********************************************************
' Adds values to arr (assumed to be an array). values can be a single value or an array.
'*********************************************************

   On Error GoTo RAISE_ERR
   
   '''''''''''''''''''''''''''''''''''''''
   ' If passed array wasn't yet initialized (i.e. we need to use ReDim only instead of
   ' ReDim Preserve.
   '''''''''''''''''''''''''''''''''''''''
   If Not iArr_IsAllocated(arr) Then
      If IsArray(values) Then
         ReDim arr(0 To UBound(values))
         Dim q As Integer
         For q = 0 To UBound(values)
            arr(q) = values(q)
         Next q
      Else:
         ReDim arr(0 To 0)
         arr(0) = values
      End If
      Exit Sub
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' If passed array was already initialized.
   '''''''''''''''''''''''''''''''''''''''
   Dim nextPositionInArray As Integer
   If IsArray(values) Then
      Dim elementsToAdd As Integer: elementsToAdd = UBound(values) + 1
      Dim j As Integer
      For j = 1 To elementsToAdd Step 1
         nextPositionInArray = UBound(arr) + 1
         ReDim Preserve arr(0 To nextPositionInArray)
         arr(nextPositionInArray) = values(j - 1)
      Next j
   Else:
      nextPositionInArray = UBound(arr) + 1
      ReDim Preserve arr(0 To nextPositionInArray)
      arr(nextPositionInArray) = values
   End If
   Exit Sub
   
RAISE_ERR:
   iError_RaiseCustom ig_ERR_ARRAYS, "iArr_AddValue"
   
End Sub

Public Function iArr_Combine1DArrays(ByVal arrayOfArrays As Variant) As Variant()
'*********************************************************
' Join multiple 1-dimensional arrays together. For example, result of adding (1,2,3,4) & (5,6)
' will be (1,2,3,4,5,6).
' This does not remove duplicates.
'*********************************************************

   On Error GoTo ERR_HANDLER
   
   Dim results() As Variant
   ReDim results(0)
   Dim nextPositionInResult As Long
   nextPositionInResult = 0
   
   Dim arrayCount As Integer
   arrayCount = UBound(arrayOfArrays) + 1
   Dim q As Integer
   For q = 0 To (arrayCount - 1)
      Dim elementCountInCurrentArray As Long
      elementCountInCurrentArray = UBound(arrayOfArrays(q)) + 1
      Dim j As Long
      For j = 0 To (elementCountInCurrentArray - 1)
         ReDim Preserve results(0 To nextPositionInResult)
         results(nextPositionInResult) = arrayOfArrays(q)(j)
         nextPositionInResult = nextPositionInResult + 1
      Next j
   Next q
   iArr_Combine1DArrays = results
   Exit Function
   
ERR_HANDLER:
   iError_RaiseCustom ig_ERR_ARRAYS, "iArr_Combine1DArrays()"

End Function

Public Function iArr_ContainsValue(ByRef arr As Variant, ByVal value As Variant) As Boolean
'*********************************************************
' Returns True if value is an existing element of arr.
'*********************************************************

   On Error GoTo RAISE_ERR
   Dim q As Integer
   For q = LBound(arr) To UBound(arr) Step 1
      If arr(q) = value Then
         iArr_ContainsValue = True
         Exit Function
      End If
   Next q
   iArr_ContainsValue = False
   Exit Function
   
RAISE_ERR:
   iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ContainsValue()"

End Function

Public Function iArr_FromRange( _
   ByVal rng As Range, Optional ByVal ignoreEmptyCells As Boolean = False) As Variant
'*********************************************************
' Returns array containing the values in a given range.
' It does not remove duplicate values.
'*********************************************************
   
   Dim results() As Variant
   Dim q As Long  ' marker for next element position in resulting array
   q = 0
   ReDim results(q)
   
   Dim cell As Range
   For Each cell In rng
      If ignoreEmptyCells Then
         ' If we want to ignore empty cells, first ensure the length of the cell value is greater
         ' than 0 before adding value to our results.
         If Len(cell.value) > 0 Then
            ReDim Preserve results(0 To q)
            results(q) = cell.value
            q = q + 1
         End If
      Else:
         ReDim Preserve results(0 To q)
         results(q) = cell.value
         q = q + 1
      End If
   Next cell
   iArr_FromRange = results

End Function

Public Function iArr_CountUnique(ByVal arr As Variant) As Integer
'*********************************************************
' Returns # of unique values in arr.
'*********************************************************

   Dim vals() As Variant
   If TypeName(arr) = "Range" Then
      vals = iArr_FromRange(arr)
   Else: vals = arr
   End If
   
   Dim uniqueVals() As Variant
   uniqueVals = iArr_RemoveDuplicates(vals)
   
   iArr_CountUnique = UBound(uniqueVals) + 1

End Function

Public Function iArr_DimensionNumber(ByVal arr As Variant) As Integer
'*********************************************************
' Returns # of dimensions in an array.
'*********************************************************

   On Error GoTo RETURN_RESULT
   Dim i As Integer
   Dim tmp As Integer
   i = 0
   Do While True:
      i = i + 1
      tmp = UBound(arr, i)
   Loop
RETURN_RESULT:
   iArr_DimensionNumber = i - 1
   
End Function

Public Sub iArr_Display(ByVal arr As Variant)
'*********************************************************
' Displays contents of an array with messagebox. Format of output:
'     (0,0) = Variable
'     (0,1) = Variable
' Can handle 1D and 2D arrays.
'*********************************************************

   On Error GoTo DISPLAY_ERROR
   
   Dim dimension As Integer
   dimension = iArr_DimensionNumber(arr)
   
   Dim msg As String, q As Long
   For q = 0 To UBound(arr)
      If dimension > 1 Then
         Dim j As Integer
         For j = 0 To (dimension - 1) Step 1
            msg = msg & vbLf & _
                  "[" & q & ", " & j & "] = " & arr(q, j)
         Next j
      Else:
         msg = msg & vbLf & _
               "[" & q & "] = " & arr(q)
      End If
   Next q
   MsgBox msg, , "Array"
   Exit Sub
   
DISPLAY_ERROR:
   MsgBox Prompt:="Encountered in iArr_Display", title:="Error"

End Sub

Public Function iArr_IsAllocated(ByRef arr As Variant) As Boolean
'*********************************************************
' Returns True if the array is allocated; False if array has not been allocated (i.e. a dynamic
' array that has not yet been sized with Redim, or a dynamic array that has been Erased).
'*********************************************************
   
   Dim n As Long
   ' If arr is not an array, return False and exit.
   If Not IsArray(arr) Then
      iArr_IsAllocated = False
      Exit Function
   End If
      
   ' If array has not been allocated, an error will occur getting the UBound of the array.
   On Error Resume Next
   n = UBound(arr, 1)
   If Err.number = 0 Then
      iArr_IsAllocated = True
   Else: iArr_IsAllocated = False
   End If

End Function

Public Function iArr_MinimumValue(ByRef arr As Variant) As Double
'*********************************************************
' Returns the minimum value found in arr.
'*********************************************************

   Dim result As Double
   result = arr(0)
   
   Dim q As Long
   For q = 1 To UBound(arr) Step 1
      If IsNumeric(arr(q)) Then
         If arr(q) < result Then
            result = arr(q)
         End If
      End If
   Next q
   iArr_MinimumValue = result

End Function

Public Function iArr_RemoveDuplicates(ByVal arr As Variant) As Variant()
'*********************************************************
' Returns a 1-dimensional Array with any duplicates removed.
' ASSUMES the user is passing it a 1-dimensional array and that the array is dimensioned from 0 to
' n (and not from 1 to n).
'*********************************************************
   
   On Error GoTo RAISE_ERR
   
   Dim results() As Variant
   ReDim results(0)
   
   Dim nextResult As Long ' marker for where we are in the new array.
   nextResult = 0
   
   Dim i As Long ' marker to loop through arr
   For i = 0 To UBound(arr) Step 1
      If Not iArr_ContainsValue(results, arr(i)) Then
         ReDim Preserve results(0 To nextResult)
         results(nextResult) = arr(i)
         nextResult = nextResult + 1
      End If
   Next i
   iArr_RemoveDuplicates = results()
   Exit Function
   
RAISE_ERR:
   Call Err.Raise(ig_ERR_ARRAYS, "iArr_RemoveDuplicates()", Err.description)
      
End Function

Public Sub iArr_RemoveEmptyAndNullElementsFromEnd(ByRef arr As Variant)
'*********************************************************
' Returns segment of array that does not contain Null or Empty elements. Loops from the end of the
' array towards the beginning until it finds a useful element (i.e. not Null or Empty) and then
' returns that portion of the array.
' If arr = [0, 4, Null, 2, 3, Null, Empty], then
' iArr_RemoveEmptyAndNullElementsFromEnd (arr) = [0, 4, Null, 2, 3]
'*********************************************************
   
   Dim lastNoneEmptyNullElementPosition As Integer
   lastNoneEmptyNullElementPosition = UBound(arr)
   
   Dim q As Integer
   For q = UBound(arr) To 0 Step -1
      lastNoneEmptyNullElementPosition = q
      If Not IsNull(arr(q)) And Not IsEmpty(arr(q)) Then
         Exit For
      End If
   Next q
   ReDim Preserve arr(0 To lastNoneEmptyNullElementPosition)

End Sub

Public Sub iArr_ReverseInPlace( _
   ByRef arr As Variant, _
   Optional LB As Long = -1, _
   Optional UB As Long = -1)
'*********************************************************
' Reverses the order of elements in arr.
' arr must be an allocated 1D array.
' To reverse the entire array, omit or set to less than 0 the LB and UB parameters.
' To reverse only part of tbe array, set LB and/or UB to the LBound and UBound of the sub array to
' be reversed.
'*********************************************************

   On Error GoTo ERR_HANDLER
         
   '''''''''''''''''''''''''''''''''''''''
   ' First we validate the input array, ensuring that it is:
   '  (1) an array.
   '  (2) a 1D array.
   '  (3) an array consisting of simple data types, not an array of objects or arrays.
   '''''''''''''''''''''''''''''''''''''''
   If Not IsArray(arr) Then
      iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ReverseInPlace", "Input parameter is not an array."
   End If
   
   Select Case iArr_DimensionNumber(arr)
      Case 0 ' empty, unallocated array.
         iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ReverseInPlace", "Input array is an empty, unallocated array."
      Case 1 ' We can reverse ONLY a single dimensional array
      Case Else ' multi-dimensional array
         iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ReverseInPlace", "Input array is multi-dimensional. Only 1D arrays accepted."
   End Select
   
   If Not iData_IsSimpleType(arr(LBound(arr))) Then
      iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ReverseInPlace", _
            "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
            "iArr_ReverseArrayInPlace can reverse only arrays of simple data types."
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Reverse elements now that the input array has been validated.
   '''''''''''''''''''''''''''''''''''''''
   Dim n As Long
   Dim Temp As Variant
   Dim Ndx As Long
   Dim Ndx2 As Long
   Dim OrigN As Long
   Dim NewN As Long
   Dim NewArr() As Variant
   
   If LB < 0 Then
      LB = LBound(arr)
   End If
   If UB < 0 Then
      UB = UBound(arr)
   End If
   
   For n = LB To (LB + ((UB - LB - 1) \ 2))
      Temp = arr(n)
      arr(n) = arr(UB - (n - LB))
      arr(UB - (n - LB)) = Temp
   Next n
   
   Exit Sub
   
ERR_HANDLER:
   If Err.number = ig_ERR_ARRAYS Then
      iError_RaiseExisting
   Else: iError_RaiseCustom ig_ERR_ARRAYS, "iArr_ReverseInPlace"
   End If

End Sub

Public Function iArr_SortInPlace( _
   ByRef InputArray As Variant, _
   Optional ByVal LB As Long = -1&, _
   Optional ByVal UB As Long = -1&, _
   Optional ByVal Descending As Boolean = False, _
   Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
   Optional ByVal NoAlerts As Boolean = False) As Boolean
'*********************************************************
' Sorts the original array in the calling procedure.
' INPUT:    Will work with either string data or numeric data.
'           You can sort only part of the array by setting the LB and UB (optional) parameters to
'           the first (LB) and last (UB) element indexes that you want to sort.
' NOTES:    By default, sort method is case INSENSTIVE (case doesn't matter: "A", "b", "C", "d").
'           To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode
'           argument to vbBinaryCompare (=0).
'           If Compare mode is omitted or is any value other than vbBinaryCompare,
'           it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'           The function returns TRUE if the array was successfully sorted or FALSE if an error
'           occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
'           displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'           If you coerce InputArray to a ByVal argument, iArr_SortInPlace will not be able
'           to reference the InputArray in the calling procedure and the array will not be sorted.
' MODIFYING CODE:
'           If you modify this code and you call "Exit Procedure", you MUST decrement the
'           RecursionLevel variable. E.g.:
'               If SomethingThatCausesAnExit Then
'                   RecursionLevel = RecursionLevel - 1
'                   Exit Function
'               End If
'*********************************************************

   Dim Temp As Variant
   Dim buffer As Variant
   Dim CurLow As Long
   Dim CurHigh As Long
   Dim CurMidpoint As Long
   Dim Ndx As Long
   Dim pCompareMode As VbCompareMethod
   
   ' Set the default result.
   iArr_SortInPlace = False
   
   ' This variable is used to determine the level of recursion  (the function calling itself).
   ' RecursionLevel is incremented when this procedure is called, either initially by a calling
   ' procedure or recursively by itself.
   ' The variable is decremented when the procedure exits. We do the input parameter validation
   ' only when RecursionLevel is 1 (when the function is called by another function, not when it
   ' is called recursively).
   Static RecursionLevel As Long
   
   ' Keep track of the recursion level -- that is, how many times the procedure has called itself.
   ' Carry out the validation routines only when this procedure is first called. Don't run the
   ' validations on a recursive call to the procedure.
   RecursionLevel = RecursionLevel + 1
   
   On Error GoTo RETURN_FALSE
   
   If RecursionLevel = 1 Then
      If Not IsArray(InputArray) Then
         If Not NoAlerts Then
            MsgBox "The InputArray parameter is not an array."
         End If
         RecursionLevel = RecursionLevel - 1
         Exit Function
      End If
      
      ' Test LB and UB. If < 0 then set to LBound and UBound of the InputArray.
      If LB < 0 Then
         LB = LBound(InputArray)
      End If
      If UB < 0 Then
         UB = UBound(InputArray)
      End If
      
      Select Case iArr_DimensionNumber(InputArray)
         ' Zero dimensions indicates an unallocated dynamic array.
         Case 0
            If Not NoAlerts Then
               MsgBox "The InputArray is an empty, unallocated array."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
         ' We ONLY sort single dimensional arrays.
         Case 1
         Case Else
            If NoAlerts = False Then
               MsgBox "The InputArray is multi-dimensional." & _
                     "iArr_SortInPlace works only on single-dimensional arrays."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
      End Select
      ' Ensure that InputArray is an array of simple data types, not other arrays or objects.
      ' This tests the data type of only the first element of InputArray. If InputArray is an
      ' array of Variants, subsequent data types may not be simple data types (e.g., they may be
      ' objects or other arrays), and this may cause iArr_SortInPlace to fail on the StrComp
      ' operation.
      If iData_IsSimpleType(InputArray(LBound(InputArray))) = False Then
         If NoAlerts = False Then
            MsgBox "InputArray is not an array of simple data types."
           RecursionLevel = RecursionLevel - 1
            Exit Function
         End If
      End If
      ' Ensure that the LB parameter is valid.
      Select Case LB
         Case Is < LBound(InputArray)
            If NoAlerts = False Then
               MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
         Case Is > UBound(InputArray)
            If NoAlerts = False Then
               MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
         Case Is > UB
            If NoAlerts = False Then
               MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
      End Select
      
      ' Ensure the UB parameter is valid.
      Select Case UB
         Case Is > UBound(InputArray)
            If NoAlerts = False Then
               MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
         Case Is < LBound(InputArray)
            If NoAlerts = False Then
               MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
         Case Is < LB
            If NoAlerts = False Then
               MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
            End If
            RecursionLevel = RecursionLevel - 1
            Exit Function
      End Select
   
      ' If UB = LB, we have nothing to sort, so get out.
      If UB = LB Then
         iArr_SortInPlace = True
         RecursionLevel = RecursionLevel - 1
         Exit Function
      End If
   
   End If ' RecursionLevel = 1
   
   ' Ensure that CompareMode is either vbBinaryCompare  or vbTextCompare. If it is neither, default to vbTextCompare.
   If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
      pCompareMode = CompareMode
   Else: pCompareMode = vbTextCompare
   End If
   
   ' Begin the actual sorting process.
   CurLow = LB
   CurHigh = UB
   
   If LB = 0 Then
      CurMidpoint = ((LB + UB) \ 2) + 1
   Else: CurMidpoint = (LB + UB) \ 2
   End If
   Temp = InputArray(CurMidpoint)
   
   Do While (CurLow <= CurHigh)
      
      Do While QSortCompare(v1:=InputArray(CurLow), v2:=Temp, CompareMode:=pCompareMode) < 0
         CurLow = CurLow + 1
         If CurLow = UB Then
            Exit Do
         End If
      Loop
      
      Do While QSortCompare(v1:=Temp, v2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
         CurHigh = CurHigh - 1
         If CurHigh = LB Then
            Exit Do
         End If
      Loop
   
      If (CurLow <= CurHigh) Then
         buffer = InputArray(CurLow)
         InputArray(CurLow) = InputArray(CurHigh)
         InputArray(CurHigh) = buffer
         CurLow = CurLow + 1
         CurHigh = CurHigh - 1
      End If
   Loop
   
   If LB < CurHigh Then
      Call iArr_SortInPlace(InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
         Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True)
   End If
   
   If CurLow < UB Then
      Call iArr_SortInPlace(InputArray:=InputArray, LB:=CurLow, UB:=UB, _
         Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True)
   End If
   
   ' If Descending is True, reverse the order of the array, but only if the recursion level is 1.
   If Descending = True Then
      If RecursionLevel = 1 Then
         Call iArr_ReverseInPlace(InputArray, LB, UB)
      End If
   End If
   
   ' We get here once all recursion is finished.
   RecursionLevel = 0
   iArr_SortInPlace = True
   Exit Function

RETURN_FALSE:
   RecursionLevel = 0
   iArr_SortInPlace = False
   
End Function

Private Function QSortCompare( _
   v1 As Variant, _
   v2 As Variant, _
   Optional CompareMode As VbCompareMethod = vbTextCompare) As Integer
'*********************************************************
' PURPOSE:  Used by iArr_SortInPlace to compare two elements.
' METHOD:   If v1 AND v2 are both numeric data types, they are converted to Doubles and compared.
'           If v1 AND v2 are both strings that contain numeric data, they are converted to Doubles
'           and compared.
'           If EITHER v1 or v2 is a string and does NOT contain numeric data, both v1 and v2 are
'           converted to Strings and compared with StrComp.
'           For text comparisons, case sensitivity is controlled by CompareMode. If this is
'           vbBinaryCompare, the result is case SENSITIVE. If this is omitted or any other value,
'           the result is case INSENSITIVE.
' RETURNS:  if v1 < v2 => -1
'           if v1 = v2 => 0
'           if v1 > v2 => 1
'*********************************************************
   
   Dim d1 As Double, d2 As Double, s1 As String, s2 As String
   
   ' Test CompareMode. Any value other than vbBinaryCompare will default to vbTextCompare.
   Dim Compare As VbCompareMethod
   If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
      Compare = CompareMode
   Else: Compare = vbTextCompare
   End If
   
   ' If either v1 or v2 is an array or an Object, raise error.
   If IsArray(v1) Or IsArray(v2) Then
      Call Err.Raise(ig_ERROR_TYPE_MISMATCH, "QSortCompare()")
   End If
   If IsObject(v1) Or IsObject(v2) Then
      Call Err.Raise(ig_ERROR_TYPE_MISMATCH, "QSortCompare()")
   End If
   
   If iData_IsSimpleNumericType(v1) Then
      If iData_IsSimpleNumericType(v2) Then
         ' If BOTH v1 and v2 are numeric data types, then convert to Doubles and do an
         ' arithmetic compare and return the result.
         d1 = CDbl(v1)
         d2 = CDbl(v2)
         If d1 = d2 Then
            QSortCompare = 0
            Exit Function
         End If
         If d1 < d2 Then
            QSortCompare = -1
            Exit Function
         End If
         If d1 > d2 Then
            QSortCompare = 1
            Exit Function
         End If
      End If
   End If
   ' Either v1 or v2 was not numeric data type. Test whether BOTH v1 AND v2 are numeric strings.
   ' If BOTH are numeric, convert to Doubles and do a arithmetic comparison.
   If IsNumeric(v1) And IsNumeric(v2) Then
      d1 = CDbl(v1)
      d2 = CDbl(v2)
      If d1 = d2 Then
         QSortCompare = 0
         Exit Function
      End If
      If d1 < d2 Then
         QSortCompare = -1
         Exit Function
      End If
      If d1 > d2 Then
         QSortCompare = 1
         Exit Function
      End If
   End If
   ' Either or both v1 and v2 was not numeric string. In this case, convert to Strings and use
   ' StrComp to compare.
   s1 = CStr(v1)
   s2 = CStr(v2)
   QSortCompare = StrComp(s1, s2, Compare)

End Function

Private Sub swapArrayDimensions( _
   ByRef arrToModify As Variant, _
   ByVal indexToModify As Integer, _
   ByVal arrToCopyFrom As Variant, _
   ByVal indexToCopy As Integer)
'*********************************************************
' Replace element in a given index with element corresponding to that index in a another array.
'*********************************************************

   Dim nD As Integer: nD = iArr_DimensionNumber(arrToModify)
   Dim x As Integer
   For x = 0 To (nD - 1) Step 1
      arrToModify(indexToModify, x) = arrToCopyFrom(indexToCopy, x)
   Next x

End Sub

'*****************************************************************************************
'*****************************************************************************************
' ARRAY METHOD TESTING
' This section of the module contains tests for several of the above methods / procedures.
'*****************************************************************************************
'*****************************************************************************************

Private Sub Test_RemoveNullElementsFromEnd()

   Dim arrayWithNullAndEmptyElements() As Variant
   ReDim arrayWithNullAndEmptyElements(0 To 4)
   arrayWithNullAndEmptyElements(0) = 1
   arrayWithNullAndEmptyElements(1) = 2
   arrayWithNullAndEmptyElements(2) = 3
   arrayWithNullAndEmptyElements(3) = Null
   
   Dim arrayWithoutEmptyOrNullElements As Variant
   arrayWithoutEmptyOrNullElements = Array(1, 2, 3)
   
   Call iArr_RemoveEmptyAndNullElementsFromEnd(arrayWithNullAndEmptyElements)
   
   Dim msg As String, header As String
   If UBound(arrayWithoutEmptyOrNullElements) <> UBound(arrayWithNullAndEmptyElements) Then
      msg = "The number of elements in the array without Empty or Null elements does not match the " & _
         "number in the array originally containing null elements that we tried to remove."
      header = "Test FAIL"
      GoTo NOTIFY_USER
   Else:
      Dim q As Integer
      For q = 0 To UBound(arrayWithoutEmptyOrNullElements)
         Dim val As Variant
         val = arrayWithoutEmptyOrNullElements(q)
         If Not iArr_ContainsValue(arrayWithNullAndEmptyElements, val) Then
            msg = "The array originally containing Empty and Null elements that we removed does not " & _
               "contain a value we expected it to (" & val & ")."
            header = "Test FAIL"
            GoTo NOTIFY_USER
         End If
      Next q
   End If
   
   msg = "(1) The number of elements is the same in the array defined without Null elements and " & _
      "the array originally containing Null elements that we removed. " & vbLf & _
      "(2) The arrays consisted of the same elements."
   header = "Test SUCCESS"
         
NOTIFY_USER:
   MsgBox msg, , header

End Sub

Private Sub Test_AddArrays()
'*********************************************************
'EXPECTATION:   It will display an array containing the numbers between 1-13.
'*********************************************************

   Dim array1() As Variant
   array1 = Array(1, 2, 3, 4, 5, 6)
   Dim array2() As Variant
   array2 = Array(7, 8, 9, 10)
   Dim array3() As Variant
   array3 = Array(11, 12, 13)
   
   Dim arrayOfArrays() As Variant
   arrayOfArrays = Array(array1, array2, array3)
   
   Dim combinationOfArrays() As Variant
   combinationOfArrays = iArr_Combine1DArrays(arrayOfArrays)
   
   Call iArr_Display(combinationOfArrays)

End Sub

Private Sub Test_NoDuplicates()
'*********************************************************
'EXPECTATION:   It will display an array that does not contain duplicate 1s or 2s.
'*********************************************************

   Dim testArray() As Variant
   testArray = Array(1, 1, 1, 1, 2, 2, 2, 4, 5, 6)
   
   Dim newArray() As Variant
   newArray = iArr_RemoveDuplicates(testArray)
   
   Call iArr_Display(newArray)

End Sub

Private Sub TestNewSort()

   Dim testArray() As Variant
   testArray = Array(1, 5, 3, 4, 2, 9)
   
   Dim vbCompare As VbCompareMethod
   vbCompare = vbBinaryCompare
   
   If iArr_SortInPlace(testArray, , , False, vbCompare) = False Then
      MsgBox "Error"
   Else: Call iArr_Display(testArray)
   End If

End Sub

Private Sub Test_SwapDimensions()
   
   Dim array1(0 To 2, 0 To 2) As Variant
   array1(0, 0) = 0
   array1(0, 1) = 1
   array1(0, 2) = 2
   array1(1, 0) = 1
   array1(1, 1) = 2
   
   Dim array2(0 To 2, 0 To 2) As Variant
   array2(0, 0) = 1
   array2(0, 1) = 2
   array2(0, 2) = 3
   array2(1, 0) = 2
   array2(1, 1) = 3
   
   Call swapArrayDimensions(array1, 0, array2, 0)
   
   MsgBox array1(0, 1)
   
End Sub

Private Sub TestReturnValuesFromRange()
   
   Dim myArray() As Variant
   myArray = iArr_FromRange(Selection)
   Call iArr_Display(myArray)
   
End Sub


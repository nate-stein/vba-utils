Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   TOOLS
' PURPOSE:  Contains various tools to automate a lot of common Excel and other tasks.
' METHODS:  CalculateCellsInRange
'           CellHasFormulaReferencingAnotherWorkbook
'           ChangeColumnVisibility
'           ClearComments
'           CondenseRange
'           ConvertColNumberToLetter
'           DeleteColumnsWithXInHeader
'           DeleteDuplicates
'           DeleteEmptyRows
'           HideColumnsWithXInhHeader
'           ExportCurrentReferencesToNewWorksheet
'           FillEmptyCellsWithValue
'           FindAndReplace
'           ForcePasteDates
'           ForcePasteDoubles
'           GetRangeProperties
'           IsDivisible
'           MsgNumberFormat
'           MsgActiveColumnNumber
'           MsgAddInLocationPath
'           RemovePageBreaks
'           RepasteValues
'           SaveStringToTextFile
'           ShowDesiredColumns
'           TrimValues
'           UnhideRows
'           NewBlankPage
'           TurnScreenUpdatingOFF
'           TurnScreenUpdatingON
'*****************************************************************************************
'*****************************************************************************************

Public Sub iTools_CalculateCellsInRange(Optional ByVal rng As Range)
'*********************************************************
' Calculates cells in rng (current Selection by default). Used if, for example, you didn't want to
' calculate an entire worksheet but only specific cells.
'*********************************************************
    If rng Is Nothing Then Set rng = Selection
    
    Dim cell As Range
    For Each cell In rng
        cell.Calculate
    Next cell
End Sub

Public Function iTools_CellHasFormulaReferencingAnotherWorkbook(ByVal cell As Range) As Boolean
   
   If Not cell.HasFormula Then
      iTools_CellHasFormulaReferencingAnotherWorkbook = False
      Exit Function
   End If
      
   If InStr(1, cell.formula, "!") > 0 Then
      iTools_CellHasFormulaReferencingAnotherWorkbook = True
   Else: iTools_CellHasFormulaReferencingAnotherWorkbook = False
   End If

End Function

Public Sub iTools_ChangeColumnVisibility( _
   ByVal makeAllVisible As Boolean, _
   Optional ByVal cols As Variant, _
   Optional ByVal colFurthestRight As Integer = 50)
'*********************************************************
' makeAllVisible:
'           If True, all columns in cols are made visible.
'           If False, all columns in cols are made hidden.
' Executes by checking whether each column up to colFurthestRight is in cols.
'*********************************************************
      
   Dim col As Integer
   For col = 1 To colFurthestRight
      If Not makeAllVisible Then
         If iArr_ContainsValue(cols, col) Then
            Columns(col).Hidden = True
         End If
      Else: Columns(col).Hidden = False
      End If
   Next col

End Sub

Public Sub iTools_ClearComments( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)
'*********************************************************
' Removes comments from every cell in rng (current Selection by default).
'*********************************************************
      
   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If
   
   Dim cell As Range
   For Each cell In rng
      cell.ClearComments
   Next cell
   
   If Not msgUser Then Exit Sub
   
   Dim msg As String
   If noRngSupplied Then
      msg = "Finished clearing comments from Selection."
   Else: msg = "Finished clearing comments from " & rng.Address & "."
   End If
   MsgBox msg, , "Done"
   
End Sub

Public Sub iTools_CondenseRange( _
   Optional ByVal rng As Range, _
   Optional ByVal column As Long)
'*********************************************************
' Condenses values in rng (current Selection by default) by looking for empty cells in the
' specified column. If the column is empty, then values are moved up in order to occupy the empty
' column.
' The "condensing" is accomplished by copying and pasting, NOT deleting.
'*********************************************************

   If rng Is Nothing Then Set rng = Selection
   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   ' If user didn't specify column to check, ensure rng only spans one column.
   If column = 0 Then
      If rngProps.ColumnCount > 1 Then
         MsgBox "No column was provided to iTools_CondenseRange and the Range spans more than one column.", , "Error"
         Exit Sub
      End If
      column = rngProps.LeftmostColumn
   End If
   
   TurnScreenUpdatingOFF
   
   Dim row As Long
   Dim exitSearchTextLoop As Boolean
   For row = rngProps.firstRow To rngProps.lastRow Step 1
      If IsEmpty(cells(row, column)) Then
         Dim emptySpace As Range ' Space where content may be moved up.
         Set emptySpace = Range(cells(row, rngProps.LeftmostColumn), cells(row, rngProps.RightmostColumn))
         '''''''''''''''''''''''''''''''''''''''
         ' Look for row where there is a value in column -- the contents in this row will be moved
         ' up to occupy the empty space.
         '''''''''''''''''''''''''''''''''''''''
         Dim foundContentsToMoveUp As Boolean: foundContentsToMoveUp = False
         Dim rowToMoveUp As Long: rowToMoveUp = row + 1
         Do
            If Not IsEmpty(cells(rowToMoveUp, column)) Then
               Dim rngToMove As Range
               Set rngToMove = Range(cells(rowToMoveUp, rngProps.LeftmostColumn), cells(rowToMoveUp, rngProps.RightmostColumn))
               rngToMove.Copy
               emptySpace.PasteSpecial xlPasteValues
               rngToMove.ClearContents
               foundContentsToMoveUp = True
            End If
            rowToMoveUp = rowToMoveUp + 1
         Loop Until (foundContentsToMoveUp) Or (rowToMoveUp > rngProps.lastRow)
      End If
   Next row
   
   TurnScreenUpdatingON
      
End Sub

Public Function iTools_ConvertColNumberToLetter(ByVal columnNumber As Long) As String

   Dim vArr
   vArr = Split(cells(1, columnNumber).Address(True, False), "$")
   iTools_ConvertColNumberToLetter = vArr(0)
   
End Function

Public Sub iTools_DeleteColumnsWithXInHeader()
'*********************************************************
' Deletes all columns that contain an X (capital or lowercase) in the first row.
'*********************************************************

   Dim rightmostCol As Integer
   rightmostCol = Range("ZZ1").End(xlToLeft).column
   
   Dim col As Integer
   For col = rightmostCol To 1 Step -1
      If Not IsEmpty(cells(1, col).value) Then
         If LCase(cells(1, col).value) = "x" Then
            Columns(col).Delete
         End If
      End If
   Next col

End Sub

Public Sub iTools_DeleteDuplicates( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range, _
   Optional ByVal column As Long)
'*********************************************************
' Delete duplicate entries in the passed-through range.
' Begins deleting duplicates from the bottom and progresses up.
'*********************************************************
   
   If rng Is Nothing Then Set rng = Selection
   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   ' If user didn't specify column to check, ensure rng only spans one column.
   If column = 0 Then
      If rngProps.ColumnCount > 1 Then
         MsgBox "No column was provided to iTools_DeleteDuplicates and the Range spans more than one column.", , "Error"
         Exit Sub
      End If
      column = rngProps.LeftmostColumn
   End If
   
   TurnScreenUpdatingOFF
   Dim deletedCount As Integer: deletedCount = 0
   Dim row As Long
   For row = rngProps.lastRow To rngProps.firstRow Step -1
      Dim val As Variant
      val = cells(row, column).Text
      Dim uncheckedRng As Range
      Set uncheckedRng = Range(cells(rngProps.firstRow, column), cells(row, column))
      
      If Application.WorksheetFunction.CountIf(uncheckedRng, val) > 1 Then
         Rows(row).Delete
         deletedCount = deletedCount + 1
      End If
   Next row
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
     
   Dim msg As String
   If deletedCount = 0 Then
      msg = "No rows duplicate values were found"
   Else: msg = deletedCount & " rows with duplicate values were deleted"
   End If
   msg = msg & " in column " & UCase(iTools_ConvertColNumberToLetter(column)) & "."
   MsgBox msg, , "Done"

End Sub

Public Sub iTools_DeleteEmptyRows( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range, _
   Optional ByVal column As Integer)
'*********************************************************
' Deletes rows within rng (current Selection by default) where there is no value in column. No
' value is needed for column unless rng spans more than one column.
'*********************************************************
   
   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If
   
   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   ' If user didn't specify column to check, ensure rng only spans one column.
   If column = 0 Then
      If rngProps.ColumnCount > 1 Then
         MsgBox "No column was provided to iTools_DeleteEmptyRows and the Range spans more than one column.", , "Error"
         Exit Sub
      End If
      column = rngProps.LeftmostColumn
   End If
   
   TurnScreenUpdatingOFF
   Dim deletedCount As Long: deletedCount = 0
   Dim row As Long
   For row = rngProps.lastRow To rngProps.firstRow Step -1
      If IsEmpty(cells(row, column)) Then
         Rows(row).Delete
         deletedCount = deletedCount + 1
      End If
   Next row
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
   
   '''''''''''''''''''''''''''''''''''''''
   ' Message displayed to the user will depend on whether or not any empty rows were actually
   ' encountered.
   '''''''''''''''''''''''''''''''''''''''
   Dim msg As String
   Select Case deletedCount
      Case 0
         msg = "There were no empty rows"
      Case 1
         msg = "1 empty row deleted"
      Case Is > 1
         msg = deletedCount & " empty rows deleted"
   End Select
   
   If noRngSupplied Then
      msg = msg & " in Selection."
   Else: msg = msg & " in " & rng.Address & "."
   End If
   MsgBox msg, , "Done"

End Sub

Public Sub iTools_HideColumnsWithXInHeader()
'*********************************************************
' Hides all columns that contain an X (capital or lowercase) in the first row.
'*********************************************************

   Dim rightmostCol As Integer
   rightmostCol = Range("ZZ1").End(xlToLeft).column
   
   Dim col As Integer
   For col = rightmostCol To 1 Step -1
      If Not IsEmpty(cells(1, col).value) Then
         If LCase(cells(1, col).value) = "x" Then
            Columns(col).Hidden = True
         End If
      End If
   Next col

End Sub

Public Sub iTools_ExportCurrentReferencesToNewWorksheet()
'*********************************************************
' Inserts Worksheet outlining all the current references in ActiveWorkbook.
'*********************************************************
 
   ActiveWorkbook.Sheets.Add
   ActiveSheet.name = "GUIDS"
   
   On Error Resume Next
   Dim n As Integer
   For n = 1 To ActiveWorkbook.VBProject.References.count
      cells(n, 1) = ActiveWorkbook.VBProject.References.Item(n).name
      cells(n, 2) = ActiveWorkbook.VBProject.References.Item(n).description
      cells(n, 3) = ActiveWorkbook.VBProject.References.Item(n).Guid
      cells(n, 4) = ActiveWorkbook.VBProject.References.Item(n).Major
      cells(n, 5) = ActiveWorkbook.VBProject.References.Item(n).Minor
      cells(n, 6) = ActiveWorkbook.VBProject.References.Item(n).fullpath
   Next n

End Sub

Public Sub iTools_FillEmptyCellsWithValue( _
   Optional ByVal rng As Range, Optional ByVal val As Variant)
'*********************************************************
' Fill in any blank cells in rng (current Selection by default) with a given value (a hyphen by
' default).
'*********************************************************
      
   If rng Is Nothing Then Set rng = Selection
   If Len(val) = 0 Then val = "-"
   
   Dim cell As Range
   For Each cell In rng
      If IsEmpty(cell) Then cell.value = val
   Next cell
      
End Sub

Public Sub iTools_FindAndReplace( _
   ByVal findTxt As String, _
   ByVal replaceTxt As String, _
   Optional wksName As Variant)
   
   If IsMissing(wksName) Then
      ActiveSheet.cells.Replace What:=findTxt, replacement:=replaceTxt, LookAt:=xlWhole, _
      SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
      ReplaceFormat:=False
   Else:
      Worksheets(wksName).cells.Replace What:=findTxt, replacement:=replaceTxt, LookAt:=xlWhole, _
      SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
      ReplaceFormat:=False
   End If
      
End Sub

Public Sub iTools_ForcePasteDates( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)

   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If

   TurnScreenUpdatingOFF
   Dim cell As Range, dt As Date
   For Each cell In rng
      dt = CDate(cell.value)
      cell.value = dt
   Next cell
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
      
   Dim msg As String
   If noRngSupplied Then
      msg = "Finished pasting values as Dates in Selection."
   Else: msg = "Finished pasting values as Dates in " & rng.Address & "."
   End If
   MsgBox msg, , "Done"
   
End Sub

Public Sub iTools_ForcePasteDoubles( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)
   
   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If
   
   TurnScreenUpdatingOFF
   Dim cell As Range, num As Double
   For Each cell In rng
      If Not IsEmpty(cell.value) Then
         If Len(cell.value) > 0 Then
            num = CDbl(cell.value)
            cell.value = num
         End If
      End If
   Next cell
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
   
   Dim msg As String
   If noRngSupplied Then
      msg = "Finished pasting values as Doubles in Selection."
   Else: msg = "Finished pasting values as Doubles in " & rng.Address & "."
   End If
   MsgBox msg, , "Done"

End Sub

Public Function iTools_GetRangeProperties(ByVal rng As Range) As IT_RangeProperties
'*********************************************************
' Returns IT_RangeProperties object with information populated based on rng.
'*********************************************************

   Dim result As IT_RangeProperties
   With result
      .LeftmostColumn = rng.Columns(1).column
      .RightmostColumn = rng.Columns.count + .LeftmostColumn - 1
      .firstRow = rng.Rows(1).row
      .lastRow = rng.Rows.count + .firstRow - 1
      .ColumnCount = .RightmostColumn - .LeftmostColumn + 1
      .RowCount = .lastRow - .firstRow + 1
   End With
   iTools_GetRangeProperties = result

End Function

Public Function iTools_IsDivisible(ByVal number As Double, ByVal divisor As Double) As Boolean
'*********************************************************
' Returns True if number is perfectly divisible by divisor; False otherwise.
'*********************************************************

   ' Multiply numbers by 10,000 first in case we are dealing with decimals (in which case Mod
   ' won't work).
   Const multiplier As Integer = 10000
   divisor = divisor * multiplier
   number = number * multiplier
   
   If (number Mod divisor) = 0 Then
      iTools_IsDivisible = True
   Else: iTools_IsDivisible = False
   End If
   
End Function

Public Function iTools_IsInteger(ByVal number As Double) As Boolean
'*********************************************************
' Returns True if number is perfectly divisible by one i.e. an integer; False otherwise.
'*********************************************************

   iTools_IsInteger = iTools_IsDivisible(number, 1)

End Function

Public Sub iTools_MsgNumberFormat()

   MsgBox ActiveCell.NumberFormat

End Sub

Public Sub iTools_MsgActiveColumnNumber()
      
   MsgBox "Column Number = " & ActiveCell.column, , "Column Number"
      
End Sub

Public Sub iTools_MsgAddInLocationPath()

   Dim addInPath As String
   addInPath = ThisWorkbook.path & Application.PathSeparator
   MsgBox addInPath

End Sub

Public Sub iTools_RemovePageBreaks()
'*********************************************************
' Removes the line breaks representing print margins from all Worksheets in ActiveWorkbook.
'*********************************************************

   Dim wks As Worksheet
   For Each wks In ActiveWorkbook.Worksheets
      wks.DisplayPageBreaks = False
   Next wks

End Sub

Public Sub iTools_RepasteValues( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)

   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If
   
   Dim cell As Range
   For Each cell In rng
      cell.value = cell.value
   Next cell
   
   If Not msgUser Then Exit Sub
   
   Dim msg As String
   If noRngSupplied Then
      msg = "Finished repasting values in Selection."
   Else: msg = "Finished repasting values in " & rng.Address & "."
   End If
   MsgBox msg, , "Done"

End Sub

Public Sub iTools_SaveStringToTextFile(ByVal txt As String, ByVal path As String)
'*********************************************************
' Saves input txt to the desired file path.
'*********************************************************
   
   ' Make sure file save path has a ".txt" extension.
   If Right(path, 4) <> ".txt" Then path = path & ".txt"
   
   ' Get an unused file number.
   Dim unusedFileNumber As Integer: unusedFileNumber = FreeFile
   
   ' Create a new file (or overwrite an existing one).
   Open path For Output As unusedFileNumber
   
   Print #unusedFileNumber, txt
   
   Close unusedFileNumber

End Sub

Public Sub iTools_ShowDesiredColumns( _
   ByRef headersOfColumnsToHide() As Variant, _
   Optional ByVal headerRow As Integer = 1, _
   Optional ByVal firstColumn As Integer = 1, _
   Optional ByVal lastColumn As Integer = 40)
'*********************************************************
' Checks the value in headerRow from firstColumn to lastColumn. If that value is contained
' within headersOfColumnsToHide, that column is made hidden.
'*********************************************************

   Dim col As Long
   For col = firstColumn To lastColumn Step 1
      Dim columnHeader As String
      columnHeader = cells(headerRow, col).value
      If iArr_ContainsValue(headersOfColumnsToHide, columnHeader) Then
         Columns(col).Hidden = True
      End If
   Next col

End Sub

Public Sub iTools_TrimValues( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)
   
   Dim noRngSupplied As Boolean
   If rng Is Nothing Then
      noRngSupplied = True
      Set rng = Selection
   End If
   
   TurnScreenUpdatingOFF
   Dim cell As Range
   For Each cell In rng
      cell.value = Trim(cell.value)
   Next cell
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
    
   Dim msg As String
   If noRngSupplied Then
      msg = "Finished trimming values in Selection."
   Else: msg = "Finished trimming values in " & rng.Address & "."
   End If
   MsgBox msg, , "Done"

End Sub

Public Sub iTools_UnhideRows( _
   Optional ByVal msgUser As Boolean = True, _
   Optional ByVal rng As Range)
'*********************************************************
' Ensures all rows spanned by rng (current Selection by default) are visible.
'*********************************************************

   If rng Is Nothing Then Set rng = Selection
   
   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   TurnScreenUpdatingOFF
   Dim unhiddenCount As Integer: unhiddenCount = 0
   Dim row As Long
   For row = rngProps.firstRow To rngProps.lastRow Step 1
      If Rows(row).Hidden Then
         Rows(row).Hidden = False
         unhiddenCount = unhiddenCount + 1
      End If
   Next row
   
   TurnScreenUpdatingON
   If Not msgUser Then Exit Sub
   
   Dim msg As String
   If unhiddenCount > 0 Then
      msg = "Successfully unhid " & unhiddenCount & " rows."
   Else: msg = "There were no hidden rows in Range."
   End If
   MsgBox msg, , "Done"
      
End Sub

Public Sub iTools_NewBlankPage()
'*********************************************************
' Inserts new Worksheet into ActiveWorkbook with special formatting for various purposes.
'*********************************************************

   On Error GoTo ERROR_HANDLER
   
   Const FONT_COLOR_INDEX As Integer = 2
   Const FONT_COLOR_AUTO As Integer = 0
   Const BACKGROUND_COLOR_MAJORITY As Integer = 2
   Const ZOOM_LEVEL As Integer = 125
   Const ROW_HEIGHT_GENERAL As Integer = 13
   Const ROW_HEIGHT_HEADER As Integer = 20
   Const FONT_SIZE As Integer = 8
   Const FONT_NAME As String = "Calibri"
   Const COLUMN_WIDTH_PERCENTAGES As Integer = 12
   Const COLUMN_WIDTH_TEXT As Integer = 20
   Const COLUMN_WIDTH_NUMBERS As Integer = 12
   Const LAST_COLUMN_TO_FORMAT As Integer = 6
   Const FIRST_ROW_TO_APPLY_ALTERNATE_COLORING As Integer = 2
   Const LAST_ROW_TO_APPLY_ALTERNATE_COLORING As Integer = 100
   
   ' Add new worksheet
   ActiveWorkbook.Worksheets.Add
   Dim wksNew As Worksheet
   Set wksNew = ActiveWorkbook.ActiveSheet
   wksNew.cells.Font.color = FONT_COLOR_AUTO
   ActiveWindow.Zoom = ZOOM_LEVEL
         
   ' Update new worksheet name
   wksNew.name = createNewNameForActiveWorksheet()

   ' Set general default cell properties for all cells in new worksheet.
   ' We'll then update the special columns and headers i.e. anything requiring special formatting.
   With wksNew.cells
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
      .Interior.ColorIndex = BACKGROUND_COLOR_MAJORITY
      .Font.Size = FONT_SIZE
      .Font.name = FONT_NAME
      .Rows.RowHeight = ROW_HEIGHT_GENERAL
   End With
      
   ' Define particular ranges that will contain particular formatting
   Dim rangeForNumbers As Range, rangeForPercentages As Range, rangeForText As Range
   Set rangeForNumbers = Union(Columns(2), Columns(6))
   Set rangeForPercentages = Union(Columns(3), Columns(5))
   Set rangeForText = Union(Columns(1), Columns(4))
      
   ' Add column headers
   Union(cells(1, 1), cells(1, 4)).value = "Text"
   Union(cells(1, 2), cells(1, 6)).value = "Numbers"
   Union(cells(1, 3), cells(1, 5)).value = "Percentages"
      
   ' Columns that will contain text values
   With rangeForText
      .cells.NumberFormat = "@"
      .cells.HorizontalAlignment = xlRight
      .cells.AddIndent = True
      .cells.IndentLevel = 1
      .ColumnWidth = COLUMN_WIDTH_TEXT
   End With
      
   ' Columns that will contain numbers
   With rangeForNumbers
      '.Cells.NumberFormat = "#,##0.00"
      .cells.NumberFormat = "+ [Blue]$#,##0.0;[Red]($#,##0.0);[Black]0;_(@_)"
      .ColumnWidth = COLUMN_WIDTH_NUMBERS
   End With
   
   ' Columns that will contain percentages
   With rangeForPercentages
      .cells.NumberFormat = "0.0##%"
      .ColumnWidth = COLUMN_WIDTH_NUMBERS
   End With
      
   ' Format range containing the column headers
   Dim columnHeaders As Range: Set columnHeaders = Range("A1:F1")
   
   With columnHeaders
      .cells.Font.Bold = True
      '.Cells.Interior.color = BACKGROUND_COLOR_BLUE
      '.Cells.Interior.ColorIndex = BACKGROUND_COLOR_INDEX
      .cells.Font.ColorIndex = FONT_COLOR_INDEX
      .Rows.RowHeight = ROW_HEIGHT_HEADER
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .AddIndent = False
   End With
   

    
   ' Apply alternate background color formatting to subset of range
   Dim alternateRowColoringRange As Range
   Set alternateRowColoringRange = Range(cells(FIRST_ROW_TO_APPLY_ALTERNATE_COLORING, 1), cells(LAST_ROW_TO_APPLY_ALTERNATE_COLORING, LAST_COLUMN_TO_FORMAT))
   Call iFormat_RowsWhiteAndBlue(alternateRowColoringRange)
      
   Exit Sub

ERROR_HANDLER:
   MsgBox "Unknown error encountered in iTools_NewBlankPage.", , "Unknown Error Encountered"
   Exit Sub
End Sub

Private Function createNewNameForActiveWorksheet() As String

   Dim newWorksheetName As String
   Dim newWorksheetID As Integer: newWorksheetID = 1
   'Make sure proposed name for new worksheet doesn't already exist.
   Do
      newWorksheetName = "New_" & newWorksheetID
      newWorksheetID = newWorksheetID + 1
   Loop Until Not iWks_NameExists(newWorksheetName)
   
   createNewNameForActiveWorksheet = newWorksheetName

End Function

Public Sub TurnScreenUpdatingOFF()

   Application.ScreenUpdating = False
   
End Sub

Public Sub TurnScreenUpdatingON()

   Application.ScreenUpdating = True
   
End Sub
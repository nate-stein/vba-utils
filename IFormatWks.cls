VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormatWks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' CLASS:    WORKSHEET FORMATTER
' PURPOSE:  Formats a worksheet containing tabular data based on personal
'           preferences by mapping preferred formats to different variable
'           types.

Private Const mFONTSIZE_HEADER As Double = 9
Private mProps As IT_RangeProperties
Private mColFormats() As IT_ColFormat ' array of IT_ColFormat
Private Const m_ZOOM As Integer = 85

Public Sub Format()

   mProps = iTools_GetRangeProperties(ActiveCell.CurrentRegion)
   defineColumnFormats
   applyColumnFormats
   adjustSpecialColumns
   formatColumnHeaders
   deleteDuplicatePlayerField
   ActiveWindow.Zoom = m_ZOOM
      
   On Error Resume Next
   If Not ActiveSheet.AutoFilterMode Then
      Selection.AutoFilter
   End If
   Range("B2").Select
   ActiveWindow.FreezePanes = True

End Sub

Private Sub formatColumnHeaders()
   Dim rng As Range
   Set rng = getColumnHeaderRng()
   iFormat_InteriorDeepBlue rng
   rng.cells.HorizontalAlignment = xlCenter
   rng.cells.Font.Size = mFONTSIZE_HEADER
   iFormat_ChangeFontColor IEWhite, rng
End Sub


Private Sub deleteDuplicatePlayerField()
'*********************************************************
' Should be last sub executed b/c it makes mProps outdated.
'*********************************************************
   
   Dim nameCount As Integer
   nameCount = 0
   Dim col As Integer
   For col = mProps.LeftmostColumn To mProps.RightmostColumn Step 1
      If cells(mProps.firstRow, col).value = "player" Then
         If nameCount = 1 Then
            Columns(col).EntireColumn.Delete
            Exit Sub
         End If
         nameCount = nameCount + 1
      End If
   Next col

End Sub

Private Sub adjustSpecialColumns()
'*********************************************************
' Applies special formatting to certain columns that we pay extra attention to
' or which would benefit from clearer formatting:
'     - player
'     - report (injuries, news)
'     - impact (injuries)
'     - returns (injuries)
' Should typically be executed last in a script.
'*********************************************************
   
   Dim playerArr As Variant
   playerArr = Array("player", "p1", "p2")
   
   Dim col As Integer
   For col = mProps.LeftmostColumn To mProps.RightmostColumn Step 1
      Dim colValue As String
      colValue = cells(mProps.firstRow, col).value
      If iArr_ContainsValue(playerArr, colValue) Then
         Call applySpecialColumnFormatting(col, 20)
      ElseIf colValue = "report" Then
         Call applySpecialColumnFormatting(col, 40)
      ElseIf colValue = "impact" Then
         Call applySpecialColumnFormatting(col, 40)
      ElseIf colValue = "returns" Then
         Call applySpecialColumnFormatting(col, 25)
      End If
   Next col

End Sub

Private Sub applySpecialColumnFormatting(col As Integer, colWidth As Integer)
'''''''''''''''''''''''''''''''''''''''
' Used by adjustSpecialColumns() only since formatting to be applied tends to
' be the same.
'''''''''''''''''''''''''''''''''''''''

   Columns(col).ColumnWidth = colWidth
   With Columns(col).cells
      .IndentLevel = 1
      .HorizontalAlignment = xlLeft
   End With

End Sub

Private Sub applyColumnFormats()

   Dim i As Integer, col As Integer
   i = 0
   For col = mProps.LeftmostColumn To mProps.RightmostColumn Step 1
      Dim colRng As Range
      Set colRng = getDataRngForCol(col)
      colRng.cells.NumberFormat = mColFormats(i).CellFormat
      If cells(mProps.firstRow, col).value = "NAME" Then
         colRng.cells.HorizontalAlignment = xlLeft
         colRng.cells.IndentLevel = 1
      Else:
         colRng.cells.HorizontalAlignment = xlCenter
      End If
      i = i + 1
   Next col

End Sub

Private Sub defineColumnFormats()
'*********************************************************
' Fills mColFormats with correct column formats based on sample of data points in the column.
'*********************************************************

   ReDim mColFormats(0 To mProps.ColumnCount - 1)
      
   Dim i As Integer, col As Integer
   i = 0
   For col = mProps.LeftmostColumn To mProps.RightmostColumn Step 1
      Dim fmt As IT_ColFormat
      fmt = iFD_CreateColFormat(getCellFormat(col), col)
      mColFormats(i) = fmt
      i = i + 1
   Next col

End Sub

Private Function getCellFormat(ByVal col As Integer) As String
'*********************************************************
' Returns text format to be applied to cell values in this column. The correct format is
' determined by sampling certain number of rows in the column and analyzing the values.
'*********************************************************
   
   On Error GoTo RETURN_TEXT
   
   ' Used to determine whether non-integer numbers represent percentages or just doubles.
   Const PERCENT_THRESHOLD As Double = 2
   
   ' Number of rows to check when digging deeper between type families.
   Dim n_rows As Integer
   n_rows = Application.WorksheetFunction.min(30, mProps.RowCount - 1)
   
   Dim row As Integer, val As Variant
   val = cells(mProps.firstRow + 1, col).value
   
    If IsDate(val) Then
      getCellFormat = "m/d/yyyy"
      Exit Function
   End If
   
   If Not IsNumeric(val) Then GoTo RETURN_TEXT
      
   '''''''''''''''''''''''''''''''''''''''
   ' Since we now know the value is numeric, we loop through the rows to determine what type of
   ' numbers are contained in this column: integers, decimals, or percentages?
   ' If not integers and the absolute value of the numbers is always < PERCENT_THRESHOLD, then we
   ' assume they reflect percentages.
   '''''''''''''''''''''''''''''''''''''''
   
   ' These flags are initialized to True but flipped to False if any value proving otherwise is
   ' encountered while looping through the n_rows.
   Dim alwaysUnderPctThreshold As Boolean: alwaysUnderPctThreshold = True
   Dim alwaysInteger As Boolean: alwaysInteger = True
   For row = mProps.firstRow + 1 To mProps.firstRow + n_rows
      val = cells(row, col).value
      If Abs(val) > PERCENT_THRESHOLD Then alwaysUnderPctThreshold = False
      If Not iTools_IsInteger(val) Then alwaysInteger = False
   Next row
   
   If alwaysInteger Then
      getCellFormat = "#,##0"
      Exit Function
   End If
   
   If alwaysUnderPctThreshold Then
      getCellFormat = "0.0%; (0.0)%"
      Exit Function
   End If
   
   ' Default to double.
   getCellFormat = "0.0"
        
   Exit Function
      
RETURN_TEXT:
   getCellFormat = "@"

End Function

Private Function getColumnHeaderRng() As Range

   Set getColumnHeaderRng = Range(cells(mProps.firstRow, mProps.LeftmostColumn), cells(mProps.firstRow, mProps.RightmostColumn))

End Function

Private Function getDataRngForCol(ByVal col As Integer) As Range
'*********************************************************
' Returns Range for the data column excluding the header row.
'*********************************************************
   
   Set getDataRngForCol = Range(cells(mProps.firstRow + 1, col), cells(mProps.lastRow, col))

End Function

Attribute VB_Name = "iFormat"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   EXCEL FORMATTING TOOLS
' NOTES:    When one of the method parameters is a Range, the Range is typically set to the
'           Selection by default if none is provided.
' METHODS:  AddFormulaDetailsIntoComments
'           AllComments
'           AsTableHeaders
'           ChangeFontColor
'           ChangeInteriorColor
'           ColorAlternateColumns
'           ColorAlternateRows
'           ColumnsWhiteAndBlue
'           FinancialModelStandards
'           FontBlue
'           InputRange
'           OutputRange
'           InteriorDarkBlue
'           InteriorDeepBlue
'           InteriorLightBlue
'           Number
'           Percentage
'           MsgSelectionFontRgbCode
'           MsgSelectionInteriorRgbCode
'           RowsWhiteAndBlue
'           TestChangingColors
'           GetRgbPropertiesForCustomColor

'''''''''''''''''''''''''''''''''''''''
' Variables related to default formatting we'd like for cell comments.
'''''''''''''''''''''''''''''''''''''''
Private Const m_DEFAULT_COMMENT_FONT_NAME As String = "Calibri"
Private Const m_DEFAULT_COMMENT_FONT_SIZE As Double = 9
Private Const m_DEFAULT_COMMENT_MAX_WIDTH As Double = 250

'''''''''''''''''''''''''''''''''''''''
' Table header formats.
'''''''''''''''''''''''''''''''''''''''
Private Const m_DFLT_FONT_SIZE_TABLE_HEADER As Double = 9

'''''''''''''''''''''''''''''''''''''''
' Enums designed to standardize colors used throughout the Add-In.
'''''''''''''''''''''''''''''''''''''''
' Used as a "Key" to obtain the RGB properties needed to replicate this color.
Public Enum IENUM_CUSTOM_COLOR
   IEProgrammingBlue = 1
   IEProgrammingClassName = 2
   IEAutomatic = 3
   IEGrey = 4
   IELightBlue = 5
   IEDarkerBlue = 6
   IEDeepBlue = 7
   IEFunctionOutputBlue = 8
   IEDarkBlue1 = 9 ' used only for alternate rows in context of IELightBlue1
   IELightBlue1 = 10
   IEWhite = 11
   IEDarkGreen = 12
   IEPurple = 13
   IELightYellow = 14
End Enum

' Stores the RGB properties of a color.
Public Type ITYPE_RGB_COLOR
   Red As Integer
   green As Integer
   blue As Integer
End Type
' Examples of favorite Excel schemes:
' Red & Blue ($): + [Blue]$#,##0.0;[Red]($#,##0.0);[Black]0;_(@_)
' Red & Blue (%): [Blue]+ #,##0.0%;[Red](#,##0.0%);[Black]0%;_(@_)
' Accounting: _($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)
' Pandas DateTime: yyyy-mm-dd hh:mm:ss
'*****************************************************************************************
'*****************************************************************************************

Public Sub iFormat_AddFormulaDetailsIntoComments()
'*********************************************************
' Sets the comments of each cell in the Selection equal to the formula (if there is one).
' When a cell already contains a comment, this code will keep that comment and place it after a
' "|" following the formula.
'*********************************************************

   Dim cell As Range
   For Each cell In Selection
      If cell.HasFormula Then
         Dim newCellComment As String: newCellComment = ""
         If cellContainsComment(cell) Then
            If commentIncludesCurrentCellFormula(cell) Then
               newCellComment = cell.Comment.Text
            Else: newCellComment = cell.formula & "|" & removeFormulaFromCellComment(cell.Comment.Text)
            End If
         Else: newCellComment = cell.formula
         End If
         cell.ClearComments
         cell.AddComment (newCellComment)
         formatComment cell.Comment
      End If
   Next cell

End Sub

Private Function commentIncludesCurrentCellFormula(ByVal cell As Range) As Boolean
'*********************************************************
' Returns True if the cell's current formula is contained within the text of the comment.
'*********************************************************

   Dim commTxt As String
   commTxt = cell.Comment.Text
   
   If InStr(commTxt, cell.formula) > 0 Then
      commentIncludesCurrentCellFormula = True
   Else: commentIncludesCurrentCellFormula = False
   End If

End Function

Private Function removeFormulaFromCellComment(ByVal cellComment As String) As String
'*********************************************************
' Returns cellComment where the formula has been removed.
' This function assumes that the formula was inserted into the Comment via methodology in
' iFormat_AddFormulaDetailsIntoComments.
'*********************************************************

   removeFormulaFromCellComment = Trim(Mid(cellComment, InStr(cellComment, "|") + 1))

End Function

Public Sub iFormat_AllComments()
   Dim c As Comment
   For Each c In ActiveSheet.comments
      formatComment c
   Next c
End Sub

Public Sub iFormat_AsTableHeaders(Optional ByVal rng As Range)
'*********************************************************
' Set range to formatting used for table headers.
'*********************************************************
    If rng Is Nothing Then Set rng = Selection
    
    ' Adjust colors.
    Dim interiorColor As IENUM_CUSTOM_COLOR, fontColor As IENUM_CUSTOM_COLOR
    interiorColor = IEDarkerBlue
    fontColor = IEWhite
    iFormat_ChangeInteriorColor interiorColor, rng
    iFormat_ChangeFontColor fontColor, rng
    
    ' Adjust cell formatting.
    With rng
        .Font.Bold = True
        .Font.Size = m_DFLT_FONT_SIZE_TABLE_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
    End With
        
    ' Add filters.
    rng.AutoFilter
End Sub

Public Sub iFormat_SelectionComments()
'****************************************
' Formats the Comments in the Selection.
'****************************************
    Dim cell As Range
    For Each cell In Selection
        If cellContainsComment(cell) Then formatComment cell.Comment
    Next cell
End Sub

Private Sub formatComment( _
   ByVal c As Comment, _
   Optional ByVal fontName As String = m_DEFAULT_COMMENT_FONT_NAME, _
   Optional ByVal fontSize As Double = m_DEFAULT_COMMENT_FONT_SIZE, _
   Optional ByVal maxWidth As Double = m_DEFAULT_COMMENT_MAX_WIDTH)
   With c
      .Shape.TextFrame.Characters.Font.name = fontName
      .Shape.TextFrame.Characters.Font.Size = fontSize
   End With
   resizeComment c, maxWidth
End Sub

Private Sub resizeComment(ByVal c As Comment, ByVal maxWidth As Double)
'*********************************************************
' If c's width exceeds the maxWidth, then this code will resize the comment by first calculating
' its area and then, keeping c's Width at maxWidth, solving for the Height. A small adjustment is
' made to the height with ADJ_FACTOR to ensure there's a small margin remaining afterwards.
' Because this method uses the area of the AutoSized comment as a proxy for how tall the comment
' should be, it doesn't perfectly set the height and can leave much empty space at the end if the
' comment contains a lot of spaces.
'*********************************************************

   Const ADJ_FACTOR As Double = 1.02
    
   With c
      .Shape.TextFrame.AutoSize = True
      If .Shape.Width > maxWidth Then
         Dim area As Double
         area = .Shape.Width * .Shape.Height
         .Shape.Width = maxWidth
         .Shape.Height = (area / maxWidth) * ADJ_FACTOR
      End If
   End With

End Sub

Public Sub iFormat_ChangeFontColor( _
    ByVal color As IENUM_CUSTOM_COLOR, Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   Dim rgbColor As ITYPE_RGB_COLOR
   rgbColor = iFormat_GetRgbPropertiesForCustomColor(color)
   rng.Font.color = RGB(rgbColor.Red, rgbColor.green, rgbColor.blue)

End Sub

Public Sub iFormat_ChangeInteriorColor( _
   ByVal color As IENUM_CUSTOM_COLOR, Optional ByVal rng As Range)
   
   If rng Is Nothing Then Set rng = Selection
   Dim rgbColor As ITYPE_RGB_COLOR
   rgbColor = iFormat_GetRgbPropertiesForCustomColor(color)
   rng.Interior.color = RGB(rgbColor.Red, rgbColor.green, rgbColor.blue)

End Sub

Public Sub iFormat_ColorAlternateColumns( _
   ByVal rng As Range, _
   ByVal color1 As IENUM_CUSTOM_COLOR, _
   ByVal color2 As IENUM_CUSTOM_COLOR)
'****************************************
' Colors alternate visible columns.
'****************************************

   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   Dim useFirstColor As Boolean
   useFirstColor = True
   Dim col As Integer
   For col = rngProps.LeftmostColumn To rngProps.RightmostColumn Step 1
      If Not Columns(col).Hidden Then
         Dim columnRng As Range
         Set columnRng = Range(cells(rngProps.firstRow, col), cells(rngProps.lastRow, col))
         If useFirstColor Then
            iFormat_ChangeInteriorColor color1, columnRng
         Else: iFormat_ChangeInteriorColor color2, columnRng
         End If
         useFirstColor = Not useFirstColor
      End If
   Next col

End Sub

Public Sub iFormat_ColorAlternateRows( _
   ByVal rng As Range, _
   ByRef color1 As IENUM_CUSTOM_COLOR, _
   ByRef color2 As IENUM_CUSTOM_COLOR)
'*********************************************************
' Shades alternate visible rows in rng according to given colors.
'*********************************************************
         
   Dim rngProps As IT_RangeProperties
   rngProps = iTools_GetRangeProperties(rng)
   
   Dim useFirstColor As Boolean
   useFirstColor = True
   Dim row As Integer
   For row = rngProps.firstRow To rngProps.lastRow Step 1
      If Not Rows(row).Hidden Then
         Dim rowRng As Range
         Set rowRng = Range(cells(row, rngProps.LeftmostColumn), cells(row, rngProps.RightmostColumn))
         If useFirstColor Then
            iFormat_ChangeInteriorColor color1, rowRng
         Else: iFormat_ChangeInteriorColor color2, rowRng
         End If
         useFirstColor = Not useFirstColor
      End If
   Next row

End Sub

Public Sub iFormat_ColumnsWhiteAndBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ColorAlternateColumns rng, IEWhite, IELightBlue

End Sub

Public Sub iFormat_FinancialModelStandards(Optional ByVal rng As Range)
'*********************************************************
' Formats selected cells to conform to financial model standards.
'           input cell -> blue font w/ yellow background
'           formula -> black font
'           link to other worksheets -> green font
'*********************************************************
   
   If rng Is Nothing Then Set rng = Selection
   Dim cell As Range
   For Each cell In rng
      convertNumberFormatDependingOnCellContents cell
      If cell.HasFormula Then
         If iTools_CellHasFormulaReferencingAnotherWorkbook(cell) Then
            iFormat_ChangeFontColor IEDarkGreen, cell
         Else:
            iFormat_ChangeFontColor IEAutomatic, cell
         End If
      Else:
         iFormat_ChangeFontColor IEProgrammingBlue, cell
         iFormat_ChangeInteriorColor IELightYellow, cell
      End If
   Next cell
            
End Sub

Private Sub convertNumberFormatDependingOnCellContents(ByRef cell As Range)
   
   Dim newFormat As String
   
   If InStr(cell.Text, "%") > 0 Then
      newFormat = "0.0%; (0.0)%"
   ElseIf IsNumeric(cell.value) Then
      newFormat = "#,##0.0_);(#,##0.0)"
   Else: newFormat = "@"
   End If
   
   cell.NumberFormat = newFormat
   
End Sub

Public Sub iFormat_FontBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   rng.Font.color = RGB(0, 0, 255)
   
End Sub

Public Sub iFormat_InputRange(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ChangeInteriorColor IEGrey, rng
   iFormat_FontBlue rng

End Sub

Public Sub iFormat_OutputRange(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ChangeInteriorColor IEFunctionOutputBlue, rng
   iFormat_ChangeFontColor IEAutomatic

End Sub

Public Sub iFormat_InteriorDarkBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ChangeInteriorColor IEDarkerBlue, rng

End Sub

Public Sub iFormat_InteriorDeepBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ChangeInteriorColor IEDeepBlue, rng

End Sub

Public Sub iFormat_InteriorLightBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ChangeInteriorColor IELightBlue, rng

End Sub

Public Function iFormat_Number( _
   ByVal n As Double, minDecimals As Integer, maxDecimals As Integer) As String

   Dim customFormat As String
   If minDecimals = 0 And (maxDecimals = 0 Or isWholeNumber(n)) Then
      customFormat = "0"
   Else:
      customFormat = createFormatWithMultipleZerosAfterDecimal(minDecimals, maxDecimals)
   End If
   
   iFormat_Number = Format(n, customFormat)

End Function

Private Function isWholeNumber(ByVal NumberToCheck As Double) As Boolean
   
   isWholeNumber = iTools_IsDivisible(NumberToCheck, 1)

End Function

Public Function iFormat_Percentage( _
   ByVal pct As Double, _
   ByVal minDecimals As Integer, _
   ByVal maxDecimals As Integer) As String
   
   Dim customFormat As String
   If minDecimals = 0 And (maxDecimals = 0 Or isAWholePercent(pct)) Then
      customFormat = "0%"
   Else:
      customFormat = createFormatWithMultipleZerosAfterDecimal(minDecimals, maxDecimals) & "%"
   End If
   iFormat_Percentage = Format(pct, customFormat)
   
End Function

Private Function isAWholePercent(ByVal n As Double) As Boolean
   
   isAWholePercent = iTools_IsDivisible(n, 0.01)

End Function

Private Function createFormatWithMultipleZerosAfterDecimal( _
   ByVal minDecimals As Integer, _
   ByVal maxDecimals As Integer) As String
'*********************************************************
' ASSUMES that we want at least one zero after the decimal. For example: 0.0 instead of 0.
'*********************************************************

   Dim result As String: result = "0."
   Dim j As Integer
   For j = 1 To minDecimals Step 1
      result = result + "0"
   Next j
   
   If maxDecimals <= minDecimals Then GoTo EXIT_SMOOTHLY
   
   ' Now incorporate the optional digits after the decimal (i.e. the pound signs in the following
   ' expression: "0.00##%"
   Dim q As Integer
   For q = 1 To (maxDecimals - minDecimals) Step 1
      result = result + "#"
   Next q
   
EXIT_SMOOTHLY:
   createFormatWithMultipleZerosAfterDecimal = result

End Function

Private Function cellContainsComment(ByVal cell As Range) As Boolean

   On Error GoTo RETURN_FALSE
   Dim comments As String
   comments = cell.Comment.Text
   cellContainsComment = True
   Exit Function
   
RETURN_FALSE:
   cellContainsComment = False

End Function

Public Sub iFormat_MsgSelectionFontRgbCode()
'****************************************
' Messages the RGB color profile for the Selection's Font Color.
'****************************************

    MsgBox getRGBColorProfile(Selection.Font.color)

End Sub

Public Sub iFormat_MsgSelectionInteriorRgbCode()
'****************************************
' Messages the RGB color profile for the Selection's Interior Color.
'****************************************
   
   MsgBox getRGBColorProfile(Selection.Interior.color)

End Sub

Private Function getRGBColorProfile(ByVal color As Variant) As String

   Dim HEXcolor As String
   HEXcolor = Right("000000" & Hex(color), 6)
   getRGBColorProfile = "RGB (" & CInt("&H" & Right(HEXcolor, 2)) & _
      ", " & CInt("&H" & Mid(HEXcolor, 3, 2)) & _
      ", " & CInt("&H" & Left(HEXcolor, 2)) & ")"

End Function

Public Sub iFormat_RowsWhiteAndBlue(Optional ByVal rng As Range)

   If rng Is Nothing Then Set rng = Selection
   iFormat_ColorAlternateRows rng, IEWhite, IELightBlue
   
End Sub

Public Sub iFormat_RowsAlternatingBlue(Optional ByVal rng As Range)
'****************************************
' Colors interior of rows alternate shades of blue.
'****************************************

   If rng Is Nothing Then Set rng = Selection
   iFormat_ColorAlternateRows rng, IEDarkBlue1, IELightBlue1
   
End Sub

Public Sub iFormat_TestChangingColors()
'****************************************
' Shades cells according to incremental color schemes in order to test changes in color.
'****************************************
      
   Const MAX_ITERATIONS As Integer = 100
   Const ROW_START As Integer = 2, COL_LEFT As Integer = 2
   
   ' Starting RGB properties.
   Const R_START As Integer = 20, G_START As Integer = 20, B_START As Integer = 20
   
   ' Amount to increment each color bucket by on each loop.
   Const R_INCR As Integer = 1, G_INCR As Integer = 1, B_INCR As Integer = 1
   
   ' Add header.
   cells(ROW_START, COL_LEFT).value = "R"
   cells(ROW_START, COL_LEFT + 1).value = "G"
   cells(ROW_START, COL_LEFT + 2).value = "B"
   cells(ROW_START, COL_LEFT + 3).value = "Color"
   
   Dim q As Integer, r As Integer, g As Integer, b As Integer
   For q = 0 To MAX_ITERATIONS Step 1
      r = R_START + (R_INCR * q)
      g = G_START + (G_INCR * q)
      b = B_START + (B_INCR * q)
      
      cells(ROW_START + q + 1, COL_LEFT).value = r
      cells(ROW_START + q + 1, COL_LEFT + 1).value = g
      cells(ROW_START + q + 1, COL_LEFT + 2).value = b
      cells(ROW_START + q + 1, COL_LEFT + 3).Interior.color = RGB(r, g, b)
      
   Next q


End Sub

Public Function iFormat_GetRgbPropertiesForCustomColor( _
   ByVal color As IENUM_CUSTOM_COLOR) As ITYPE_RGB_COLOR
'****************************************
' Returns RGB properties for given color.
'****************************************

   Dim r As Integer
   Dim g As Integer
   Dim b As Integer
   
   Select Case color
      Case IEProgrammingBlue
         r = 0
         g = 0
         b = 255
      Case IEProgrammingClassName
         r = 43
         g = 145
         b = 175
      Case IEAutomatic
         r = 0
         g = 0
         b = 0
      Case IEGrey
         r = 217
         g = 217
         b = 217
      Case IELightBlue
         r = 223
         g = 237
         b = 245
      Case IEDarkerBlue
         r = 93
         g = 139
         b = 189
      Case IEDeepBlue
         r = 68
         g = 84
         b = 106
      Case IEFunctionOutputBlue
         r = 136
         g = 181
         b = 236
      Case IEDarkBlue1
         r = 174
         g = 197
         b = 218
      Case IELightBlue1
         r = 226
         g = 220
         b = 250
      Case IEWhite
         r = 255
         g = 255
         b = 255
      Case IEPurple
         r = 76
         g = 0
         b = 153
      Case IEDarkGreen
         r = 0
         g = 153
         b = 0
      Case IELightYellow
         r = 255
         g = 255
         b = 153
   End Select
   
   iFormat_GetRgbPropertiesForCustomColor.Red = r
   iFormat_GetRgbPropertiesForCustomColor.green = g
   iFormat_GetRgbPropertiesForCustomColor.blue = b

End Function


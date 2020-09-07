Attribute VB_Name = "iCommandBar"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   COMMANDBAR
' PURPOSE:  Installs the Add-In menu onto the Excel CommandBar.

Private Const m_COMMANDBAR_MENU_NAME As String = "Trinity"
Private m_AddInMenu As CommandBarControl
'*****************************************************************************************
'*****************************************************************************************

Public Sub iCommandBar_Install()
'*********************************************************
' Calls all necessary methods to update the Add-Ins CommandBar with a menu
' containing links to all desired macros.
'*********************************************************
    ' Delete existing Menu Item so that we don't have conflicting CommandBars.
    iCommandBar_RemoveAddInFromMenu
    
    On Error GoTo errHandler
    insertAddInToMenu
    addFormattingMacros
    addSSMacros
    addUtilsMacros
    Exit Sub
      
errHandler:
    Dim msg As String
    msg = "Error encountered while attempting to add the Trinity Add-In menu." & vbLf & _
        "Error Number: " & Err.number & vbLf & _
        "Error Description: " & Err.description
    MsgBox msg, , "Error"
End Sub

Private Sub addSSMacros()
    Dim submenu As CommandBarControl
    Set submenu = createMenuItem("SafeSpace")
    
    '''''''''''''''''''''''''''''''''''''''
    ' Worksheet formatting
    '''''''''''''''''''''''''''''''''''''''
    Dim fmt As CommandBarControl
    Set fmt = createMenuItem("Format", submenu)
    addMacroToMenu fmt, "Comments", "iSSFmt_Comments"
    addMacroToMenu fmt, "Posts", "iSSFmt_Posts"
End Sub

Private Sub addFDMacros()

   Dim submenu As CommandBarControl
   Set submenu = createMenuItem("FD")
   
   '''''''''''''''''''''''''''''''''''''''
   ' Live Exercise macros.
   '''''''''''''''''''''''''''''''''''''''
   Dim liveExc As CommandBarControl
   Set liveExc = createMenuItem("Live Exercise", submenu)
   addMacroToMenu liveExc, "Update Data", "iFD_UpdateLiveExercise"
   addMacroToMenu liveExc, "Update Column Visibility", "iFD_HideUnhideLiveExerciseColumns"
   addMacroToMenu liveExc, "Add Vegas Odds", "iFD_AddVegasOddsToLiveExercise"
   
   '''''''''''''''''''''''''''''''''''''''
   ' Worksheet formatting.
   '''''''''''''''''''''''''''''''''''''''
   Dim formatCtrl As CommandBarControl
   Set formatCtrl = createMenuItem("Format Wks", submenu)
   addMacroToMenu formatCtrl, "Single", "iFD_FormatWks"
   addMacroToMenu formatCtrl, "All", "iFD_FormatAllWks"
   
   '''''''''''''''''''''''''''''''''''''''
   ' Cell adjustment.
   '''''''''''''''''''''''''''''''''''''''
   Dim cellCtrl As CommandBarControl
   Set cellCtrl = createMenuItem("Adjust Cells", submenu)
   addMacroToMenu cellCtrl, "Divide", "iFD_DivideCells"
   addMacroToMenu cellCtrl, "Highlight Large", "iFD_HighlightLargeCells"
   addMacroToMenu cellCtrl, "Increment", "iFD_IncrementCells"

End Sub

Private Sub addFormattingMacros()
    Dim submenu As CommandBarControl
    Set submenu = createMenuItem("Format")
    
    '''''''''''''''''''''''''''''''''''''''
    ' Formatting Range according to pre-defined templates.
    '''''''''''''''''''''''''''''''''''''''
    Dim asMenu As CommandBarControl
    Set asMenu = createMenuItem("As", submenu)
    addMacroToMenu asMenu, "Table Headers", "iFormat_AsTableHeaders"
    addMacroToMenu asMenu, "Input", "iFormat_InputRange", True
    addMacroToMenu asMenu, "Output", "iFormat_OutputRange"
    
    '''''''''''''''''''''''''''''''''''''''
    ' Color menu item for macros only concerned with changing colors of selection.
    '''''''''''''''''''''''''''''''''''''''
    Dim colorMenu As CommandBarControl
    Set colorMenu = createMenuItem("Color", submenu)
    
    ' Interior color submenu
    Dim interiorMenu As CommandBarControl
    Set interiorMenu = createMenuItem("Interior", colorMenu)
    addMacroToMenu interiorMenu, "Dark Blue", "iFormat_InteriorDarkBlue"
    addMacroToMenu interiorMenu, "Deep Blue", "iFormat_InteriorDeepBlue"
    addMacroToMenu interiorMenu, "Light Blue", "iFormat_InteriorLightBlue"
    
    
    ' Color rows alternating colors
    Dim rowColorsMenu As CommandBarControl
    Set rowColorsMenu = createMenuItem("Alternate Rows", colorMenu)
    addMacroToMenu rowColorsMenu, "White and Blue", "iFormat_RowsWhiteAndBlue"
    addMacroToMenu rowColorsMenu, "Alternating Blues", "iFormat_RowsAlternatingBlue"
    
    ' Rest of color macros
    addMacroToMenu colorMenu, "Columns White and Blue", "iFormat_ColumnsWhiteAndBlue", True
    addMacroToMenu colorMenu, "Font Blue", "iFormat_FontBlue"
    
    '''''''''''''''''''''''''''''''''''''''
    ' Macros only concerned with comments.
    '''''''''''''''''''''''''''''''''''''''
    Dim commentsMenu As CommandBarControl
    Set commentsMenu = createMenuItem("Comments", submenu)
    addMacroToMenu commentsMenu, "Clear from Selection", "iFormat_AddFormulaDetailsIntoComments"
    addMacroToMenu commentsMenu, "Insert Formula Details", "iFormat_AddFormulaDetailsIntoComments"
    addMacroToMenu commentsMenu, "Format All", "iFormat_AllComments"
    addMacroToMenu commentsMenu, "Format Selection", "iFormat_SelectionComments"
    
    '''''''''''''''''''''''''''''''''''''''
    ' Messaging RGB codes.
    '''''''''''''''''''''''''''''''''''''''
    Dim rgbMenu As CommandBarControl
    Set rgbMenu = createMenuItem("Message RGB Code", submenu)
    addMacroToMenu rgbMenu, "Font", "iFormat_MsgSelectionFontRgbCode"
    addMacroToMenu rgbMenu, "Interior Color", "iFormat_MsgSelectionInteriorRgbCode"

End Sub

Private Sub addUtilsMacros()

   Dim submenu As CommandBarControl
   Set submenu = createMenuItem("Utils")
   
   '''''''''''''''''''''''''''''''''''''''
   ' Submenu for macros only concerned with pasting values in the Selection as a certain type.
   '''''''''''''''''''''''''''''''''''''''
   Dim forcePasteMenu As CommandBarControl
   Set forcePasteMenu = createMenuItem("Force Paste As", submenu)
   addMacroToMenu forcePasteMenu, "Dates", "iTools_ForcePasteDates"
   addMacroToMenu forcePasteMenu, "Doubles", "iTools_ForcePasteDoubles"
   addMacroToMenu forcePasteMenu, "Trimmed Values", "iTools_TrimValues"
   addMacroToMenu forcePasteMenu, "Existing Values", "iTools_RepasteValues"
   
   '''''''''''''''''''''''''''''''''''''''
   ' Macros concerned with deleting values in the Selection as a certain type.
   '''''''''''''''''''''''''''''''''''''''
   Dim deleteMenu As CommandBarControl
   Set deleteMenu = createMenuItem("Delete", submenu)
   addMacroToMenu deleteMenu, "Duplicates", "iTools_DeleteDuplicates"
   addMacroToMenu deleteMenu, "Empty Rows", "iTools_DeleteEmptyRows"
   addMacroToMenu deleteMenu, "Columns with X in Header", "iTools_DeleteColumnsWithXInHeader"
   
   '''''''''''''''''''''''''''''''''''''''
   ' General utils.
   '''''''''''''''''''''''''''''''''''''''
   addMacroToMenu submenu, "Calculate Selection", "iTools_CalculateCellsInRange", True
   addMacroToMenu submenu, "Condense Selection", "iTools_CondenseRange"
   addMacroToMenu submenu, "Convert to Standard Casing", "iText_ConvertToStandardCase"
   addMacroToMenu submenu, "Fill Empty Cells with Hyphen", "iTools_FillEmptyCellsWithValue"
   addMacroToMenu submenu, "Hide Columns with X in Header", "iTools_HideColumnsWithXInHeader"
   addMacroToMenu submenu, "Insert Blank Page", "iTools_NewBlankPage"
   addMacroToMenu submenu, "Msg Column Number", "iTools_MsgActiveColumnNumber"
   addMacroToMenu submenu, "Msg Cell Format", "iTools_MsgNumberFormat"
   addMacroToMenu submenu, "Unhide Rows", "iTools_UnhideRows"
    
End Sub

Private Function createMenuItem( _
   ByVal caption As String, _
   Optional ByRef subMenuItem As CommandBarControl) As CommandBarControl
'*********************************************************
' Creates and returns a new submenu item with a given caption under the main Add-In menu.
' If you want a new control to be nested within an existing menu item, you can pass that menu item
' as subMenuItem.
'*********************************************************

   Dim newItem As CommandBarControl
   If subMenuItem Is Nothing Then
      Set newItem = m_AddInMenu.Controls.Add(Type:=msoControlPopup)
   Else: Set newItem = subMenuItem.Controls.Add(Type:=msoControlPopup)
   End If
   newItem.caption = caption
   Set createMenuItem = newItem
   Set newItem = Nothing

End Function

Private Sub addMacroToMenu( _
   ByRef menuItem As CommandBarControl, _
   ByVal macroCaption As String, _
   ByVal macroName As String, _
   Optional ByVal insertSeparatorBefore As Boolean = False)
'*********************************************************
' Adds new macro menu item under menuItem (i.e. menuItem is the parent item for the new macro).
' macroCaption:   Display name for this new macro in the Add-In menu.
' macroName:      Actual name of macro to execute when user clicks this option.
' insertSeparatorBefore: If True, then a soft line will appear above this macro item in the menu;
'                 otherwise, no divider will be displayed.
'*********************************************************
   
   Dim newControlButton As CommandBarControl
   Set newControlButton = menuItem.Controls.Add(Type:=msoControlButton)
   
   newControlButton.caption = macroCaption
   newControlButton.OnAction = macroName
   newControlButton.BeginGroup = insertSeparatorBefore

End Sub

Public Sub iCommandBar_RemoveAddInFromMenu()
'*********************************************************
' Removes the Add-In from the Menu.
'*********************************************************
      
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(m_COMMANDBAR_MENU_NAME).Delete
      
End Sub

Private Sub insertAddInToMenu()

    ' Declare and define new Menu item in the CommandBar.
    Dim mainMenu As CommandBar
    Set mainMenu = Application.CommandBars("Worksheet Menu Bar")
    ' Set new Menu Item's location to the last location on the CommandBar
    Dim location As Integer
    location = mainMenu.Controls.count + 1
    Set m_AddInMenu = mainMenu.Controls.Add(Type:=msoControlPopup, Before:=location)
    m_AddInMenu.caption = m_COMMANDBAR_MENU_NAME

End Sub

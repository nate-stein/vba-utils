Attribute VB_Name = "iWks"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   WORKSHEET UTILS
' FUNCS:    ColumnLetterFromNumber
'           IsVisible
'           NameExists
'*****************************************************************************************
'*****************************************************************************************

Public Function iWks_ColumnLetterFromNumber(ByVal columnNumber As Long) As String
'*********************************************************
' Returns the column letter corresponding to a given column number.
' For example: iWks_ColumnLetterFromNumber(5) = "E"
'*********************************************************

   Dim vArr
   vArr = Split(cells(1, columnNumber).Address(True, False), "$")
   iWks_ColumnLetterFromNumber = vArr(0)

End Function

Public Function iWks_IsVisible(ByVal wksName As String) As Boolean
'*********************************************************
' Returns True if the worksheet name passed by caller is visible in the ActiveWorkbook.
' Returns False if it's hidden or it doesn't exist.
'*********************************************************

   On Error GoTo RETURN_FALSE
   
   Dim testWks As Worksheet
   Set testWks = ActiveWorkbook.Worksheets(wksName)
   iWks_IsVisible = testWks.Visible
   Exit Function
   
RETURN_FALSE:
   iWks_IsVisible = False

End Function

Public Function iWks_NameExists(ByVal wksName As String) As Boolean
'*********************************************************
' Returns True if wksName is the name of a worksheet in the ActiveWorkbook.
'*********************************************************
      
   On Error GoTo RETURN_FALSE
   
   Dim testWks As Worksheet
   Set testWks = ActiveWorkbook.Worksheets(wksName)
   iWks_NameExists = True
   Exit Function
   
RETURN_FALSE:
   iWks_NameExists = False
      
End Function



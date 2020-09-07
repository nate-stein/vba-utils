Attribute VB_Name = "iDeclarations"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   GLOBAL DECLARATIONS
' PURPOSE:  Declares and defines variables that are used throughout IceMan.

'''''''''''''''''''''''''''''''''''''''
' Documentation standards
'''''''''''''''''''''''''''''''''''''''
' Variable definitions in comments:
'           Should continue indended at Col 13.
' Function declarations:
'           If a function declaration exceeds Col 100, it should begin at the next line.
'           If the entire declaration still breaches Col 100 when it's on its own line, then each
'           line should only contain one variable.
'*****************************************************************************************
'*****************************************************************************************

'''''''''''''''''''''''''''''''''''''''
' VBA Error constants
'''''''''''''''''''''''''''''''''''''''
Public Const ig_ERROR_TYPE_MISMATCH As Integer = 13

'''''''''''''''''''''''''''''''''''''''
' Custom Error constants
'''''''''''''''''''''''''''''''''''''''
Public Const ig_ERR_ARRAYS As Integer = 2005
Public Const ig_ERR_DATA_WRAPPER As Integer = 2006
Public Const ig_ERR_PY As Integer = 2008

'''''''''''''''''''''''''''''''''''''''
' Excel VBA constants.
'''''''''''''''''''''''''''''''''''''''
Public Const ig_HORIZ_ALIGN_LEFT As Long = -4131
Public Const ig_HORIZ_ALIGN_CENTER As Long = -4108


'''''''''''''''''''''''''''''''''''''''
' Wrapper for properties of a Range.
'''''''''''''''''''''''''''''''''''''''
Public Type IT_RangeProperties
   firstRow As Long
   lastRow As Long
   LeftmostColumn As Long
   RightmostColumn As Long
   RowCount As Long
   ColumnCount As Long
End Type

Public Type IT_ColFormat
   ColNumber As Integer
   CellFormat As String
   Alignment As Long
   Indent As Integer
End Type

'*****************************************************************************************
'*****************************************************************************************
' FanDuel-specific

'''''''''''''''''''''''''''''''''''''''
' Type to store information on Vegas odds.
'''''''''''''''''''''''''''''''''''''''
Public Type IT_VegasOdds
   TeamName As String
   OpenPoints As Double
   OpenSpread As Double
End Type

'''''''''''''''''''''''''''''''''''''''
' Universal error code used when raising errors to indicate that it is an error that was handled.
'''''''''''''''''''''''''''''''''''''''
Public Const ig_ERROR_FANDUEL As Integer = 3006

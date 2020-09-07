Attribute VB_Name = "iRegex"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   REGULAR EXPRESSION UTILS
' PURPOSE:  Methods to utilize VBScript.RegExp objects.
' METHODS:  GetMatches
'           GetMatchValuesArray
'*****************************************************************************************
'*****************************************************************************************

Public Function iRegex_GetMatches(ByVal txt As String, ByVal pattern As String) As Object
'*********************************************************
' Returns a Matches Collection that contains a Match Object for each match found in txt.
' RegExp.Execute() returns an empty Matches Collection if no match is found.
'*********************************************************

    Dim regx As Object
    Set regx = CreateObject("VBScript.RegExp")
    With regx
        .MultiLine = False
        .Global = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    Set iRegex_GetMatches = regx.Execute(txt)
    Set regx = Nothing

End Function

Public Function iRegex_GetMatchValuesArray( _
   ByVal txt As String, _
   ByVal pattern As String, _
   Optional ByVal returnNullIfNoMatches As Boolean = True) As Variant
'*********************************************************
' Returns array of matching values extracted from a Matches Collection.
'*********************************************************
   
   Dim matches As Object
   Set matches = iRegex_GetMatches(txt, pattern)
   If IsEmpty(matches) And returnNullIfNoMatches Then
      iRegex_GetMatchValuesArray = Null
      Exit Function
   End If
   
   Dim q As Integer: q = 0
   Dim results() As Variant
   ReDim results(0 To matches.count - 1)
   Dim match As Variant
   For Each match In matches
      results(q) = match.value
      q = q + 1
   Next match
   iRegex_GetMatchValuesArray = results

End Function

Attribute VB_Name = "iText"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   TEXT UTILS
' PURPOSE:  Methods to retrieve information about and manipulate text.
' METHODS:  ContainsALetter
'           ConvertASCIICodesToString
'           ConvertToASCIICodes
'           ConvertToCharArray
'           ConvertToStandardCase
'           IsALetter
'           IsUCaseLetter
'*****************************************************************************************
'*****************************************************************************************

Public Function iText_RemoveNonAlphaNumeric(ByVal s As String) As String
    Dim i As Integer
    Dim result As String
    
    For i = 1 To Len(s)
        Select Case Asc(Mid(s, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                result = result & Mid(s, i, 1)
        End Select
    Next
    iText_RemoveNonAlphaNumeric = result
End Function

Public Function iText_ContainsALetter(ByVal txt As String) As Boolean
'*********************************************************
' Returns True if any of the characters in the string are a letter (whether uppercase or
' lowercase).
'*********************************************************

   Dim chrs() As String
   ReDim chrs(0 To Len(txt) - 1)
   chrs = iText_ConvertToChars(txt)
   
   Dim q As Integer
   For q = 0 To UBound(chrs)
      If iText_IsALetter(chrs(q)) Then
         iText_ContainsALetter = True
         Exit Function
      End If
   Next q
   
   iText_ContainsALetter = False

End Function

Public Function iText_ConvertASCIICodesToString(ByVal codes As Variant) As String
'********************************************************
' Returns string formed by concatenating all characters in codes, an array of ASCII codes.
'********************************************************
   
   '''''''''''''''''''''''''''''''''''''''
   ' Handle cases where codes is not an array or only contains one element.
   '''''''''''''''''''''''''''''''''''''''
   If Not IsArray(codes) Then
      iText_ConvertASCIICodesToString = Chr(codes)
      Exit Function
   ElseIf UBound(codes) = 0 Then
      iText_ConvertASCIICodesToString = Chr(codes(0))
      Exit Function
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' Getting here means codes is an array containing > 1 element.
   '''''''''''''''''''''''''''''''''''''''
   Dim result As String
   result = Chr(codes(0))
   
   Dim q As Integer
   For q = 1 To UBound(codes) Step 1
      result = result & Chr(codes(q))
   Next q
      
   iText_ConvertASCIICodesToString = result

End Function

Public Function iText_ConvertToASCIICodes(ByVal txt As String) As Variant
'********************************************************
' Returns array of ASCII codes for characters in txt.
'********************************************************
   
   Dim chrs As Variant
   chrs = iText_ConvertToChars(txt)
   
   Dim result() As Variant
   ReDim result(0 To UBound(chrs))
   
   Dim i As Integer
   For i = 0 To UBound(chrs)
      result(i) = Asc(chrs(i))
   Next i
   
   iText_ConvertToASCIICodes = result

End Function

Public Function iText_ConvertToChars(ByVal txt As String) As Variant
'********************************************************
' Returns array of chars that make up txt.
'********************************************************

    Dim chrs() As String
    ReDim chrs(Len(txt) - 1)
    Dim i As Integer
    For i = 1 To Len(txt)
        chrs(i - 1) = Mid$(txt, i, 1)
    Next
    iText_ConvertToChars = chrs

End Function

Public Sub iText_ConvertToStandardCase(Optional ByVal rng As Range)
'********************************************************
' Converts text like "UPPERCASE EXAMPLE" to "Uppercase Example" where only the first letters of
' words are capitalized.
'********************************************************

   If rng Is Nothing Then Set rng = Selection
   
   Dim cell As Range
   For Each cell In rng
      Dim replacementText As String
      replacementText = convertToStandardCase(Trim(cell.Text))
      cell.value = replacementText
   Next cell

End Sub

Private Function convertToStandardCase(ByVal txt As String) As String
'*********************************************************
' Returns txt formatted in such a way that only the first letter of its word(s) is capitalized.
' For example, convertToStandardCase("MY FRIEND") = "My Friend"
'*********************************************************

   Dim words() As String
   words = Split(LCase(txt))
   
   ' We always capitalize the first word in the txt.
   words(0) = capitalizeFirstLetter(words(0))
   Dim q As Integer
   For q = 1 To UBound(words)
      If Not wordShouldBeLCase(words(q)) Then
         words(q) = capitalizeFirstLetter(words(q))
      End If
   Next q
   
   convertToStandardCase = concatenateWordsIntoExpression(words)

End Function

Private Function concatenateWordsIntoExpression(ByRef words() As String) As String
'*********************************************************
' Returns the String formed by concatenating all the elements in words array with a space in
' between them.
'*********************************************************

   Dim result As String
   Dim q As Integer
   For q = 0 To UBound(words) Step 1
      If q = 0 Then
         result = words(q)
      Else: result = result + " " + words(q)
      End If
   Next q
   concatenateWordsIntoExpression = result

End Function

Private Function capitalizeFirstLetter(ByVal word As String) As String

   Dim firstLetter As String, remainderOfWord As String
   firstLetter = Left(word, 1)
   remainderOfWord = Mid(word, 2, Len(word) - 1)
   
   capitalizeFirstLetter = UCase(firstLetter) + remainderOfWord

End Function

Private Function wordShouldBeLCase(ByVal word As String) As Boolean
'*********************************************************
' Returns True if word is a word that should generally be lowercase.
'*********************************************************

   word = LCase(word)
   
   If firstCharacterIsALetter(word) Then
      Dim lcaseWords() As Variant
      lcaseWords = getWordsThatShouldBeLCase()
      wordShouldBeLCase = iArr_ContainsValue(lcaseWords, word)
   Else: wordShouldBeLCase = True
   End If

End Function

Private Function getWordsThatShouldBeLCase() As Variant

   getWordsThatShouldBeLCase = Array("a", "an", "and", "of", "on", "the")

End Function

Private Function firstCharacterIsALetter(ByVal word As String) As Boolean

   If iText_IsALetter(Left(word, 1)) Then
      firstCharacterIsALetter = True
   Else: firstCharacterIsALetter = False
   End If

End Function

Public Function iText_IsALetter(ByVal c As String) As Boolean
'*********************************************************
' Returns True if c is a letter (whether uppercase or lowercase).
'*********************************************************

   On Error GoTo RETURN_FALSE
   Select Case Asc(c)
      Case 65 To 90, 97 To 122 ' ASCII Code for A to Z and a to z.
         iText_IsALetter = True
      Case Else
         iText_IsALetter = False
   End Select
   Exit Function
   
RETURN_FALSE:
   iText_IsALetter = False

End Function

Public Function iText_IsUCaseLetter(ByVal c As String) As Boolean
'*********************************************************
' Returns True if c is an uppercase letter.
'*********************************************************
   
   Select Case Asc(c)
      Case 65 To 90 ' ASCII Code for A to Z.
         iText_IsUCaseLetter = True
      Case Else
         iText_IsUCaseLetter = False
   End Select

End Function

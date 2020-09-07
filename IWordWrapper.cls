VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' CLASS:    MICROSOFT WORD APPLICATION WRAPPER
' PURPOSE:  Provide functionality around creating and editing Microsoft Word documents.

' Microsoft Word Objects
Private mWordApp As Object   ' Word Application
Private mWordDoc As Object   ' Word document

' Constants
Private Const m_FINDREPLACE_LENGTH_LIMIT As Integer = 255
Private Const m_WORD_FILE_SUFFIX_2003 As String = ".doc"
Private Const m_WORD_FILE_SUFFIX_2007 As String = ".docx"
Private Const m_WORD_FILE_SUFFIX_MACROENABLED_2007 As String = ".docm"

' Error Numbers
Private Const m_ERRORNUMBER_FILENOTFOUND As Integer = 5174

' Word Enums
Private Const m_WORD_FINDWRAP_ENUM_CONTINUE As Integer = 1
'*****************************************************************************************
'*****************************************************************************************

Public Sub OpenDocument(ByVal filePath As String, ByVal readOnly As Boolean)
'*********************************************************
'PURPOSE:   Open a Word document.
'INPUTS:    Complete file path of Word document we want to open.
'           Boolean for whether we would like to open the requested document as ReadOnly.
'*********************************************************
      
    On Error Resume Next
    
    defineWordApp
    
    Set mWordDoc = mWordApp.Documents.Open(filePath, , readOnly)
    
    ' Make sure we are able to find the requested document.
    If Err.number = m_ERRORNUMBER_FILENOTFOUND Then
        MsgBox "No file was found at the below location," & vbLf & vbLf & filePath, , "Error in OpenDocument()"
        Call closeWordApp
        Exit Sub
    Else:
        If mWordDoc Is Nothing Then
            MsgBox "Error Number: " & Err.number & vbLf & _
            "Error Description: " & Err.description, , "Unknown Error Encountered in OpenDocument()"
            Call closeWordApp
            Exit Sub
        End If
    End If
      
End Sub

Public Sub CloseDocument(ByVal fileSavePath As String, ByVal newFileName As String)
'*********************************************************
'PURPOSE:   Close and properly save document.
'INPUTS:    Destination Folder (without the final '\'
'           File Name we want to save new document with (without file extension).
'ASSUMPT:   New document is being saved as a Word 2003-compatible document.
'*********************************************************
      
    On Error GoTo ERROR_HANDLER
    
    ' Check whether the directory to where we want to save the folder exists; if it doesn't, we will create a new folder for it.
    If Not iFolder_PathExists(fileSavePath) Then MkDir (fileSavePath)
    
    ' Ensure file name contains Word extension.
    Dim cleanedFileName As String
    If Not fileNameContainsWordTypeExtension(newFileName) Then
        cleanedFileName = fileSavePath & "\" & addWordFileExtension(newFileName)
    Else: cleanedFileName = fileSavePath & "\" & newFileName
    End If
    
    Dim completeFilePath As String, proposedCompleteFilePath As String
    proposedCompleteFilePath = fileSavePath & "\" & addWordFileExtension(newFileName)
    
    ' Check that no previous version of this document was saved.
    If iFolder_PathExists(proposedCompleteFilePath) Then
        Dim numberToAddEndOfFile As Integer
        Do
            numberToAddEndOfFile = numberToAddEndOfFile + 1
            completeFilePath = fileSavePath & "\" & newFileName & "V" & numberToAddEndOfFile & m_WORD_FILE_SUFFIX_2003
        Loop Until iFolder_PathExists(completeFilePath) = False
        MsgBox "A document with the proposed file name and folder path already existed. A new version was saved as V" & numberToAddEndOfFile, , "Prior Version of Document"
    Else: completeFilePath = proposedCompleteFilePath
    End If
    
    mWordDoc.SaveAs (completeFilePath)
    mWordDoc.Close
    
    Call closeWordApp
    
    Exit Sub
    
ERROR_HANDLER:
    
    Dim messg As String
    messg = "Unknown error encountered in CloseDocument." & vbLf & _
        "Error Number: " & Err.number & vbLf & _
        "Description: " & Err.description
    
    MsgBox messg, , "Error Encountered"
    
    Call closeWordApp
      
End Sub

Private Function fileNameContainsWordTypeExtension(ByVal fileName As String) As Boolean

    Dim myWordExtArray() As Variant: myWordExtArray = getWordFileExtensionArray()
    Dim j As Integer
    For j = 0 To UBound(myWordExtArray())
        ' Determine whether the string contains the given Word file extension.
        If InStr(1, fileName, myWordExtArray(j)) Then
            ' Determine whether the Word file extension is at the end of the document.
            Dim extensionLength As Integer   ' length of the Word extension returned
            extensionLength = Len(myWordExtArray(j))
            If Right(fileName, extensionLength) = myWordExtArray(j) Then
                fileNameContainsWordTypeExtension = True
                Exit Function
            End If
        End If
    Next j
    
    fileNameContainsWordTypeExtension = False

End Function

Private Function addWordFileExtension(ByVal FileNameWithoutExtension) As String

    addWordFileExtension = FileNameWithoutExtension & m_WORD_FILE_SUFFIX_2003

End Function

Private Sub defineWordApp()
'*********************************************************
'PURPOSE:   Define Word Application through late-binding.
'*********************************************************
      
    On Error GoTo ERROR_MANAGER
    
    Set mWordApp = CreateObject("Word.Application")
    mWordApp.Visible = False
    Exit Sub
    
ERROR_MANAGER:
    MsgBox "Microsoft Word was not opened correctly so program was aborted before opening file.", , "Error in OpenDocument()"
    MsgBox "Error Number: " & Err.number & vbLf & _
    "Description: " & Err.description
      
End Sub

Private Function getWordFileExtensionArray() As Variant()
'*********************************************************
'PURPOSE:   Return an array representing different Word document file extensions.
'*********************************************************
      
    getWordFileExtensionArray = Array(m_WORD_FILE_SUFFIX_2003, m_WORD_FILE_SUFFIX_2007, m_WORD_FILE_SUFFIX_MACROENABLED_2007)
      
End Function

Private Sub closeWordApp()
'*********************************************************
'PURPOSE:   Release memory from Word Application object.
'*********************************************************
      
    ' If we have created a Word Application object, quite the application (saving changes) and then release its memory.
    If Not mWordApp Is Nothing Then
        ' Close Word Application only if there are no open documents (if there are, it means an existing instance of
        ' Word was grabbed and we don't want to close Word Application).
        If mWordApp.Documents.count = 0 Then
            mWordApp.Quit (True)
        End If
        Set mWordApp = Nothing
    End If

End Sub

Public Function FindText(ByVal TextToFind As String) As Boolean
'*********************************************************
'PURPOSE:   Use Word's Find feature to find a given text. Returns TRUE if the given text is found and FALSE if it wasn't.
'INPUTS:    String of text we want to find in mWordDoc.
'*********************************************************

    mWordApp.Selection.HomeKey Unit:=6 'ensure the search starts at the beginning of the document (this is equivalent to doing CTRL + HOME when working in Word)
    mWordApp.Selection.Find.replacement.ClearFormatting 'ensure the search doesn't take into account any formatting matches
    
    With mWordApp.Selection.Find
        .Text = TextToFind
        .Forward = True
        .Wrap = 1 'wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        If .Found = False Then
            FindText = False
        Else: FindText = True
        End If
    End With

End Function

Public Sub TypeText(ByVal TextToType As String)

    mWordApp.Selection.TypeText TextToType

End Sub

Public Function FindAndReplace(ByVal TextToFind As String, ByVal replacementText As String, _
    ByVal ReplaceAll As Boolean) As Boolean
'*********************************************************
'PURPOSE:   Find text in a Word document and replace it with desired text.
'INPUTS:    Text to find.
'           Text to replace found text with.
'           Boolean over whether to replace all instances of the text we're looking for or just the first we encounter.
'NOTES:     Maximum length of text to substitute with (and find) is 255 characters. So if the substituting text is
'           longer than that, we use the TypeText method.
'           Word replacement setting early/late binding references:
'               wdReplaceOne = 1
'               wdReplaceAll = 2
'               wdStory = 6
'*********************************************************

    On Error GoTo ERROR_TRAP
    
    ' Define replacement setting (1 to replace one instance; 2 to replace all).
    Dim wdReplacementSetting As Integer
    If ReplaceAll = True Then
        wdReplacementSetting = 2
    Else: wdReplacementSetting = 1
   End If
    
    ' Ensure search starts at beginning of document (equivalent to CTRL + HOME when working in Word).
    mWordApp.Selection.HomeKey Unit:=6
    
    With mWordApp.Selection.Find
        .replacement.ClearFormatting
        .Text = TextToFind
        .Forward = True
        .Wrap = m_WORD_FINDWRAP_ENUM_CONTINUE
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If Len(replacementText) > m_FINDREPLACE_LENGTH_LIMIT Then
            .Execute
            If .Found = True Then mWordApp.Selection.TypeText replacementText
        Else:
            .replacement.Text = replacementText
            .Execute Replace:=wdReplacementSetting
        End If
        ' Return True if we found the passed-through text or false if we didn't.
        If .Found = False Then
            FindAndReplace = False
        Else: FindAndReplace = True
        End If
    End With
    
    Exit Function
    
ERROR_TRAP:
    MsgBox "Error Number: " & Err.number & vbLf & _
    "Error Description: " & Err.description, , "Error in FindAndReplace"

End Function

Public Sub RemoveLine(ByVal TextToFind As String)
'*********************************************************
'PURPOSE:   Remove a line in Word document by finding a string of text and typing two backspaces.
'ASSUMPT:   Only two backspaces are required to remove the requested line (extra spaces may throw off this code).
'NOTES:     Finds the given text and then types two backspaces once text is selected.
'*********************************************************

    Call FindText(TextToFind)
    With mWordApp.Selection
        .TypeBackspace
        .TypeBackspace
    End With

End Sub

Public Function DeleteRow(ByVal TextToFind As String) As Boolean
'*********************************************************
'PURPOSE:   Delete a row containing string of text passed through to function
'ASSUMPT:   String we are looking for is contained in a table. If it is not, program will delete much more than we would have wanted.
'*********************************************************

    If FindText(TextToFind) = True Then
        mWordApp.Selection.Rows.Delete
        DeleteRow = True
    Else: DeleteRow = False
    End If

End Function

Public Function RemoveText(ByVal TextToFind As String) As Boolean
'*********************************************************
'PURPOSE:   Find and remove the string passed through to procedure
'ASSUMPT:   The only reason we are unable to remove text (and return True) is that the Find function doesn't find the _
            desired text.
'*********************************************************

   If FindText(TextToFind) Then
        mWordApp.Selection.TypeBackspace
        RemoveText = True
    Else: RemoveText = False
    End If

End Function

Public Function IsAtStartOfPage(ByVal vtxtToAnalyze As String) As Boolean
'*********************************************************
'PURPOSE:   Determine whether the text passed through has any room from its top edge to the the top of the page.
'*********************************************************

    If mWordApp.Selection.Information(8) <> 0 Then
        IsAtStartOfPage = False
   Else: IsAtStartOfPage = True
    End If

End Function

Attribute VB_Name = "iFolder"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   FOLDER UTILS
' PURPOSE:  Functions and methods used to retrive information about and create capabilities to
'           manipulate folders.
' METHODS:  ChangeFileNamesInFolder
'           CreateNew
'           DeleteFile
'           GetFileExtension
'           GetFilesCollection
'           PathExists
'*****************************************************************************************
'*****************************************************************************************

Public Sub iFolder_ChangeFileNamesInFolder()
'*********************************************************
' Finds all files matching a certain Regex pattern and changes the file name.
' This was based on previous exercise and would have to be updated.
'*********************************************************
   
   Const FOLDER_PATH As String = "C:\Users\"
   Const FILE_REGEX_PATTERN As String = "\d{4}-(\d{4}).*xlsx"
   
   Dim filesCollection As New Collection
   Set filesCollection = iFolder_GetFilesCollection(FOLDER_PATH, "*.xlsx")
   Dim q As Integer
   For q = 1 To filesCollection.count Step 1
      Dim newFileName As String
      newFileName = iRegex_GetMatches(filesCollection(q), FILE_REGEX_PATTERN).Item(0).SubMatches(0) & ".xlsx"
      Dim newFilePath As String
      newFilePath = FOLDER_PATH & newFileName
      Name filesCollection(q) As newFilePath
   Next q

End Sub

Public Function iFolder_CreateNew( _
   ByVal path As String, Optional ByVal msgUserIfPathAlreadyExists As Boolean = False) As Boolean
'*********************************************************
' Returns True if a new path was created and False if given path already existed.
'*********************************************************

   If Not iFolder_PathExists(path) Then
      MkDir (path)
      iFolder_CreateNew = True
      Exit Function
   End If
   
   iFolder_CreateNew = False
   If msgUserIfPathAlreadyExists Then
      MsgBox "File path that already exists was passed to iFolder_CreateNew()." & vbLf & _
      "Passed file path: " & path, , "FYI"
   End If

End Function

Public Sub iFolder_DeleteFile(ByVal path As String)
'*********************************************************
' Deletes file located at path.
'*********************************************************

   If Not iFolder_PathExists(path) Then Exit Sub
   
   ' Remove readonly attribute, if set
   SetAttr path, vbNormal
   ' Then delete the file
   Kill path

End Sub

Public Function iFolder_GetFileExtension(ByVal path) As String
      
   iFolder_GetFileExtension = Right(path, Len(path) - InStrRev(path, "."))
      
End Function

Public Function iFolder_GetFilesCollection( _
   ByVal folderPath As String, Optional ByVal fileRegexPattern As String = "") As Collection
   
   Dim results As New Collection
   If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
   
   ' Dir() returns first file matching criteria in folder.
   Dim file As Variant
   file = Dir(folderPath & fileRegexPattern)
   Do While Len(file) > 0
      results.Add folderPath & file
      file = Dir() ' get next file in the same folder
   Loop
   Set iFolder_GetFilesCollection = results

End Function

Public Function iFolder_PathExists(ByVal path As String) As Boolean
'*********************************************************
' Returns True if a file or folder with same directory as path exists.
'*********************************************************

   If Len(Dir(path, vbDirectory)) = 0 Then
      iFolder_PathExists = False
   Else: iFolder_PathExists = True
   End If

End Function

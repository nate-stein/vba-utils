Attribute VB_Name = "iError"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   ERROR UTILS
' PURPOSE:  Improve error-handling throughout Add-In.
' METHODS:  RaiseCustom
'           RaiseExisting
'*****************************************************************************************
'*****************************************************************************************

Public Sub iError_RaiseCustom( _
   ByVal errNumber As Integer, _
   ByVal method As String, _
   Optional ByVal descriptionOverride As String = "")
'*********************************************************
' Raises a custom Err object.
' descriptionOverride:
'           If none is provided, then the description that will be included in raised error is the
'           description from VBA's Err.Description. This is done because if the user chooses to
'           raise an custom error somewhere along the code (where an error wasn't being raised by
'           VBA itself), there should be a reason that can be provided.
'*********************************************************

   Dim description As String
   If Len(descriptionOverride) > 0 Then
      description = descriptionOverride
   Else: description = Err.description
   End If
   
   Err.Raise errNumber, method, description

End Sub

Public Sub iError_RaiseExisting( _
   Optional ByVal additionalSource As String = "", _
   Optional ByVal additionalDetails As String = "")
'*********************************************************
' Raises existing error by taking existing error's information. User can pass additional
' information.
'*********************************************************

   Dim source As String
   source = Err.source
   If Len(additionalSource) > 0 Then source = source & " | " & additionalSource
   
   Dim description As String
   description = Err.description
   If Len(additionalDetails) > 0 Then description = description & " | " & additionalDetails
   
   Err.Raise Err.number, source, description

End Sub


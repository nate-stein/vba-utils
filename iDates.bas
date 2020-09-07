Attribute VB_Name = "iDates"
Option Explicit
'*****************************************************************************************
'*****************************************************************************************
' MODULE:   DATE UTILS
' METHODS:  Format
'           GetFancyDay
'           ParseFromText
'*****************************************************************************************
'*****************************************************************************************

Public Function iDates_Format(ByVal dt As Date, ByVal formatRef As Integer) As String
'*********************************************************
' Returns a date that is formatted according to formatRef passed by caller.
' Format codes:
'           1: 14-Feb-2001
'           2: February 14, 2001
'           3: February 14th, 2001
'           4: next Wednesday (14-Feb) ** the "smart" description **
'           5: 14-Feb
'           6: 2/14/2001
'           7: 010214
'*********************************************************
   
   Select Case formatRef
      Case 1
         iDates_Format = Format(dt, "D-MMM-YYYY")
      Case 2
         iDates_Format = Format(dt, "MMMM D, YYYY")
      Case 3
         iDates_Format = Format(dt, "MMMM") & " " & iDates_GetFancyDay(dt) & ", " & Format(dt, "YYYY")
      Case 4
         iDates_Format = createSmartDateDescription(dt)
      Case 5
         iDates_Format = Format(dt, "D-MMM")
      Case 6
         iDates_Format = Format(dt, "M/D/YYYY")
      Case 7
         Dim yr As Integer, m As Integer, d As Integer
         yr = Year(dt)
         m = Month(dt)
         d = Day(dt)
         iDates_Format = Right(yr, 2) & Format(m, "00") & Format(d, "00")
      Case Else
         iDates_Format = "Unrecognized formatRef."
   End Select

End Function

Private Function createSmartDateDescription( _
   ByVal dt As Date, Optional ByVal dateFormatRef As Integer = 5) As String
'*********************************************************
' Returns a "smart" description of dt.
' dateFormatRef: Used to format the date portion of the result.
'*********************************************************

   ' How we format the date itself is the same regardless of scenario. What will vary depending
   ' on where in the week we are with dt is the context that is provided with it.
   Dim formattedDate As String
   formattedDate = iDates_Format(dt, dateFormatRef)
   
   If dt = Date Then
      createSmartDateDescription = "today " & "(" & formattedDate & ")"
      Exit Function
   ElseIf (dt - Date) = 1 Then
      createSmartDateDescription = "tomorrow " & "(" & formattedDate & ")"
      Exit Function
   End If
   
   '''''''''''''''''''''''''''''''''''''''
   ' If dt is neither today or tomorrow, more logic goes into how we format dt.
   '''''''''''''''''''''''''''''''''''''''
   Dim weeksAfterToday As Integer
   weeksAfterToday = weeksDateTakesPlaceAfterToday(dt)
   
   ' No 'smart' annotations are made if dt takes place more than 1 week after today.
   If weeksAfterToday >= 2 Then
      createSmartDateDescription = "on " & formattedDate
      Exit Function
   End If
   
   ' A 'smart' prefix is added to dt if it takes place this week or next week.
   Dim introTxt
   If weeksAfterToday = 0 Then
      introTxt = "this"
   ElseIf weeksAfterToday = 1 Then
      If (Weekday(Date, vbMonday) - Weekday(dt, vbMonday)) <= 0 Then
         introTxt = "next"
      Else:
         introTxt = "this upcoming"
      End If
   End If
   
   createSmartDateDescription = introTxt & " " & Format(dt, "DDDD") & " (" & formattedDate & ")"

End Function

Public Function weeksDateTakesPlaceAfterToday(ByVal dt As Date) As Integer
'*********************************************************
' Returns the number of weeks away from Today that dt takes place.
' For example, if Today were Friday, 12/30/2016, then would return:
'        0 if dt took place this week i.e. b/t Friday, 12/30/2016 & Sunday, 1/1/2017 (inclusive).
'        1 if dt took place b/t Monday, 1/2/2017 & Sunday, 1/8/2017.
'        2 if dt took place b/t Monday, 1/9/2017 & Sunday, 1/15/2017.
'*********************************************************
   
   Dim totalDays As Integer, daysAfterThisWeek As Integer
   totalDays = (dt - Date)
   daysAfterThisWeek = totalDays - getDaysRemainingThisWeek()
   
   If daysAfterThisWeek < 1 Then
      weeksDateTakesPlaceAfterToday = 0
   Else:
      weeksDateTakesPlaceAfterToday = Application.WorksheetFunction.RoundDown((daysAfterThisWeek - 1) / 7, 0) + 1
   End If
   
End Function

Private Function getDaysRemainingThisWeek()
'*********************************************************
' Returns number of days that will pass before the upcoming Monday (exclusive of Today).
' For example, if dt were equal to Friday 12/30/2016, this would return 2 (= 7 - 5).
'*********************************************************

   getDaysRemainingThisWeek = 7 - Weekday(Date, vbMonday)

End Function

Public Function iDates_GetFancyDay(ByVal dt As Date) As String
      
      Dim fanyDescriptions(1 To 31) As String
      fanyDescriptions(1) = "1st"
      fanyDescriptions(2) = "2nd"
      fanyDescriptions(3) = "3rd"
      fanyDescriptions(4) = "4th"
      fanyDescriptions(5) = "5th"
      fanyDescriptions(6) = "6th"
      fanyDescriptions(7) = "7th"
      fanyDescriptions(8) = "8th"
      fanyDescriptions(9) = "9th"
      fanyDescriptions(10) = "10th"
      fanyDescriptions(11) = "11th"
      fanyDescriptions(12) = "12th"
      fanyDescriptions(13) = "13th"
      fanyDescriptions(14) = "14th"
      fanyDescriptions(15) = "15th"
      fanyDescriptions(16) = "16th"
      fanyDescriptions(17) = "17th"
      fanyDescriptions(18) = "18th"
      fanyDescriptions(19) = "19th"
      fanyDescriptions(20) = "20th"
      fanyDescriptions(21) = "21st"
      fanyDescriptions(22) = "22nd"
      fanyDescriptions(23) = "23rd"
      fanyDescriptions(24) = "24th"
      fanyDescriptions(25) = "25th"
      fanyDescriptions(26) = "26th"
      fanyDescriptions(27) = "27th"
      fanyDescriptions(28) = "28th"
      fanyDescriptions(29) = "29th"
      fanyDescriptions(30) = "30th"
      fanyDescriptions(31) = "31st"
      iDates_GetFancyDay = fanyDescriptions(Day(dt))

End Function

Public Function iDates_ParseFromText(ByVal dt As String, ByVal formatRef As Integer) As Variant
'*********************************************************
' Format refs guide us as to how we should parse the date from a String (i.e. the date string is
' of the following form) :
'        1: 20160220
'*********************************************************

   Dim yr As Integer, m As Integer, d As Integer
   Select Case formatRef
      Case 1
         yr = Left(dt, 4)
         m = Mid(dt, 5, 2)
         d = Right(dt, 2)
         Dim result As Date: result = m & "/" & d & "/" & yr
         iDates_ParseFromText = result
      Case Else
         iDates_ParseFromText = "Unhandled format ref (" & formatRef & ") encountered."
   End Select

End Function

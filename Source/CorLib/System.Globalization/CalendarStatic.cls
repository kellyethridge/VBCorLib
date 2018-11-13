VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: CalendarStatic
'

''
' Provides static methods for the <a href="Calendar.html">Calendar</a> class.
'
' @remarks To access these methods, use the variable name <b>Calendar</b> directly.
' <pre>
' Era = Calendar.CurrentEra
' </pre>
'
Option Explicit

Private mCachedDateTime As New CorDateTime


''
' Returns the current era for any calendar class.
'
' @return The current era.
' @remarks A constant is returned with the value of 0.
' @see Calendar
'
Public Property Get CurrentEra() As Long
    CurrentEra = 0
End Property

''
' Returns a read-only version of the specified <a href="Calendar.html">Calendar</a> object.
'
' @param Calendar The <a href="Calendar.html">Calendar</a> object.
' @return The <a href="Calendar.html">Calendar</a> object specified by the calendar parameter, if calendar is read-only.<br/>
' -or-<br/>
' A read-only memberwise clone of the <a href="Calendar.html">Calendar</a> object specified by <i>calendar</i>, if <i>calendar</i> is not read-only.
'
Public Function ReadOnly(ByVal Calendar As Calendar) As Calendar
    If Calendar Is Nothing Then _
        Error.ArgumentNull "Calendar"
    
    If Calendar.IsReadOnly Then
        Set ReadOnly = Calendar
    Else
        If TypeOf Calendar Is HebrewCalendar Then
            Set ReadOnly = HebrewCalendar.ReadOnly(Calendar)
        ElseIf TypeOf Calendar Is GregorianCalendar Then
            Set ReadOnly = GregorianCalendar.ReadOnly(Calendar)
        ElseIf TypeOf Calendar Is JulianCalendar Then
            Set ReadOnly = JulianCalendar.ReadOnly(Calendar)
        ElseIf TypeOf Calendar Is HijriCalendar Then
            Set ReadOnly = HijriCalendar.ReadOnly(Calendar)
        End If
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Common base Calendar functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetWeekOfYear(ByVal Calendar As Calendar, ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    Dim FirstWeekLength As Long
    Dim Offset          As Long
    Dim WholeWeeks      As Long
    Dim doy             As Long

    FirstWeekLength = FirstDayOfWeek - Calendar.GetDayOfWeek(Calendar.ToDateTime(Calendar.GetYear(Time), 1, 1, 0, 0, 0, 0))
    
    If FirstWeekLength < 0 Then
        FirstWeekLength = FirstWeekLength + 7
    End If

    Select Case Rule
        Case FirstDay
            If FirstWeekLength > 0 Then
                Offset = 1
            End If
            
        Case FirstFullWeek
            If FirstWeekLength >= 7 Then
                Offset = 1
            End If
            
        Case FirstFourDayWeek
            If FirstWeekLength >= 4 Then
                Offset = 1
            End If
    End Select

    doy = Calendar.GetDayOfYear(Time)
    
    If doy > FirstWeekLength Then
        WholeWeeks = (doy - FirstWeekLength) \ 7
        
        If WholeWeeks * 7 + FirstWeekLength < doy Then
            Offset = Offset + 1
        End If
    End If

    GetWeekOfYear = WholeWeeks + Offset
    
    If GetWeekOfYear = 0 Then
        Dim Year    As Long
        Dim Month   As Long
        Dim Day     As Long
        
        Year = Calendar.GetYear(Time) - 1
        Month = Calendar.GetMonthsInYear(Year)
        Day = Calendar.GetDaysInMonth(Year, Month)
        GetWeekOfYear = GetWeekOfYear(Calendar, Calendar.ToDateTime(Year, Month, Day, 0, 0, 0, 0), Rule, FirstDayOfWeek)
    End If
End Function

Friend Function ToFourDigitYear(ByVal Calendar As Calendar, ByVal Year As Long) As Long
    Dim y As Long
    
    If Year < 0 Then _
        Error.ArgumentOutOfRange "Year", ArgumentOutOfRange_NeedNonNegNum
    
    If Year < 100 Then
        y = Year Mod 100
        ToFourDigitYear = (Calendar.TwoDigitYearMax \ 100) * 100 + y
        
        If y > Calendar.TwoDigitYearMax Mod 100 Then
            ToFourDigitYear = ToFourDigitYear - 100
        End If
    Else
        ToFourDigitYear = Year
    End If
End Function

Friend Sub CheckAddResult(ByVal Milliseconds As Currency, ByVal MinValue As CorDateTime, ByVal MaxValue As CorDateTime)
    If Milliseconds < MinValue.TotalMilliseconds Or Milliseconds > MaxValue.TotalMilliseconds Then
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ResultCalendarRange, MinValue, MaxValue))
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetCalendarInfoInt32(ByVal Cal As Long, ByVal CalType As Long) As Long
    If GetCalendarInfo(LOCALE_USER_DEFAULT, Cal, CalType Or CAL_RETURN_NUMBER, vbNullString, 0, GetCalendarInfoInt32) = BOOL_FALSE Then
        Error.Win32Error Err.LastDllError
    End If
End Function



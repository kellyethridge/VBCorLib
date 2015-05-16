VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTimeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: cDateTimeStatic
'

''
' Represents methods to create cDateTime objects from various data sources.
'
' @remarks These methods are considered static. The <b>cDateTimeStatic</b> cannot
' be directly instantiated. Use these methods withe <b>cDateTime.*</b> syntax.
' <pre>
' Dim dt As cDateTime
' Set dt = cDateTime.FromOADate(#1/1/2001 8:30:00 AM#)
'
' ' A quicker way to create a <b>cDateTime</b> object is to use the <b>NewDate</b> method.
' Set dt = NewDate(#1/1/2001 8:30:00 AM#)
' </pre>
' This will create a new instance of <b>cDateTime</b> with the supplied date value.
' @see cDateTime
'
Option Explicit

Private Const SECONDS_PER_MINUTE        As Currency = 60@
Private Const MINUTES_PER_HOUR          As Currency = 60@
Private Const HOURS_PER_DAY             As Long = 24

Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private Const DAYS_TO_18991231          As Long = 693593
Private Const MILLISECONDS_TO_18991231  As Currency = 1@ * DAYS_TO_18991231 * MILLISECONDS_PER_DAY

' we put this here instead of cDateTime to prevent conflicts with
' the cDateTime.DayOfWeek property.

''
' This enumeration represents the days of the week as used by cDateTime and Calendars.
'
' @param Sunday 0
' @param Monday 1
' @param Tuesday 2
' @param Wednesday 3
' @param Thursday 4
' @param Friday 5
' @param Saturday 6
' @remarks The DayOfWeek is 0 based, where as VB's vbDayOfWeek is 1 based.
Public Enum DayOfWeek
    Sunday
    Monday
    Tuesday
    Wednesday
    Thursday
    Friday
    Saturday
End Enum

''
' Used to specify the type of cDateTime object.
'
' @param UnspecifiedKind 0
' @param UtcKind 1
' @param LocalKind 2
'
Public Enum DateTimeKind
    UnspecifiedKind = 0
    UtcKind = 1
    LocalKind = 2
End Enum


Private mDaysInMonth()      As Long
Private mDateAscDateTime    As cDateTime
Private mMinValue           As cDateTime
Private mMaxValue           As cDateTime


''
' Returns a boolean indicating if the specific Gregorian year is a leap year.
'
' @param Year The year to test for leap year.
' @return Indication of leap year.
' @remarks Only the Gregorian calendar is supported through the
' cDateTime functions. To determine leap years for other calendar
' types, an instance of the specific calendar will need to be used.
' <p>If the year is evenly divisible by 4 or 400 and not by 100
' then it is a leap year.</p>
'
Public Function IsLeapYear(ByVal Year As Long) As Boolean
    If (Year Mod 4) = 0 Then
        IsLeapYear = (Year Mod 100) <> 0 Or (Year Mod 400) = 0
    End If
End Function

''
' Returns the number of days in a Gregorian calendar month.
'
' @param Year The year the month is in.
' @param Month The month to find the number of days for.
' @return The number of days in the specified month.
'
Public Function DaysInMonth(ByVal Year As Long, ByVal Month As Long) As Long
    If Year < 1 Or Year > 9999 Then _
        Error.ArgumentOutOfRange "Year", ArgumentOutOfRange_Year
    If Month < 1 Or Month > 12 Then _
        Error.ArgumentOutOfRange "Month", ArgumentOutOfRange_Month
    
    If Month = 2 Then
        If IsLeapYear(Year) Then
            DaysInMonth = 29
            Exit Function
        End If
    End If
    DaysInMonth = mDaysInMonth(Month - 1)
End Function

''
' Returns a cDateTime object from a VB Date value.
'
' @param d The standard VB Date value.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return The cDateTime representation of the date.
'
Public Function FromOADate(ByVal d As Date, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set FromOADate = New cDateTime
    FromOADate.InitFromDate d, Kind
End Function

''
' Creates a <b>cDateTime</b> object from an operating system file timestamp.
'
' @param time The file time in UTC time.
' @return The cDateTime of the file time converted to local time from UTC time.
' @remarks Time is a 64-bit integer representing a Windows file timestamp. Using
' a Currency datatype means the timestamp is represented as milliseconds with the
' fraction representing ticks of a millisecond. No conversion if performed. The
' 64-bit values is used as is.
' <p>The timestamp is the number of milliseconds (including the fractional tick portion)
' elasped since January 1, 1601 A.D. 12:00 A.M. coordinated universal time (UTC).</p>
' <p>A tick is 100-nanoseconds. There are 10000 ticks per millisecond, equivalent to
' the number of fractional values in a Currency datatype.</p>
'
Public Function FromFileTime(ByVal Time As Currency) As cDateTime
    Set FromFileTime = FromFileTimeUtc(Time).ToLocalTime
End Function

''
' Returns a cDateTime object from a FileTime value.
'
' @param time The file time in UTC time.
' @return The cDateTime of the file time in UTC time.
'
Public Function FromFileTimeUtc(ByVal Time As Currency) As cDateTime
    If Time < 0 Then _
        Error.NegativeNumber "Time"
    
    Set FromFileTimeUtc = New cDateTime
    FromFileTimeUtc.InitFromFileTime Time, UtcKind
End Function

''
' Returns a cDateTime object from milliseconds.
'
' @param ms The number of milliseconds the time is equal to.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return The cDateTime object converted from milliseconds.
' @remarks Since a Currency can be passed in, the fractional portion
' can be used to represent the number ticks (1/10000 ms) in the date.
'
Public Function FromMilliseconds(ByVal ms As Currency, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set FromMilliseconds = New cDateTime
    FromMilliseconds.InitFromMilliseconds ms, Kind
End Function

''
' Returns a cDateTime that represents the current time on the local machine.
'
' @return The cDateTime object representing the current time on the local machine.
' @remarks This method uses GetLocalTime which has a resolution down to the
' milliseconds. If the milliseconds is not wanted, then either the Millisecond
' must be subtracted for the instance, or use cDateTime.FromOADate(Now).
'
Public Function Now() As cDateTime
    Dim LocalTime As SYSTEMTIME
    GetLocalTime LocalTime
    With LocalTime
        Set Now = Cor.NewDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds, Kind:=LocalKind)
    End With
End Function

''
' Returns a cDateTime that represents the current time with the UTC offset.
'
' @return The cDateTime object representing the current time with the UTC offset.
' @remarks This method uses GetSystemTime which has a resolution down to the
' milliseconds. There is no direct VB equivalent.
'
Public Function UtcNow() As cDateTime
    Dim UtcTime As SYSTEMTIME
    GetSystemTime UtcTime
    With UtcTime
        Set UtcNow = Cor.NewDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds, Kind:=UtcKind)
    End With
End Function

''
' Returns a cDateTime that represents only the current date. The time
' portion is not present.
'
' @return The cDateTime representing the date.
' @remarks The time is still there, just is only 12:00AM.
'
Public Function Today() As cDateTime
    Dim LocalTime As SYSTEMTIME
    GetLocalTime LocalTime
    With LocalTime
        Set Today = Cor.NewDateTime(.wYear, .wMonth, .wDay, Kind:=LocalKind)
    End With
End Function

''
' Returns the minimum cDateTime representation.
'
' @return The minimum cDateTime.
' @remarks The minimum cDateTime is 1/1/0001 12:00:00 AM.
'
Public Property Get MinValue() As cDateTime
    If mMinValue Is Nothing Then
        Set mMinValue = Cor.NewDateTime(1, 1, 1)
    End If
    
    Set MinValue = mMinValue
End Property

''
' Returns the maximum cDateTime representation.
'
' @return The maximum cDateTime.
' @remarks The maximum cDateTime is "12/31/9999 11:59:59 PM".
'
Public Property Get MaxValue() As cDateTime
    If mMaxValue Is Nothing Then
        Set mMaxValue = Cor.NewDateTime(9999, 12, 31, 23, 59, 59, 999)
    End If
    
    Set MaxValue = mMaxValue
End Property

''
' Parses a string representation of a date into a cDateTime object.
'
' @param s The date string to be parsed.
' @return A cDateTime object created from the date string.
' @remarks The parse method simply uses CDate to parse the string.
'
Public Function Parse(ByRef s As String) As cDateTime
    Set Parse = FromOADate(CDate(s))
End Function

''
' Attempts to parse a string representation of a date into a cDateTime object.
'
' @param s The date string to be parsed.
' @param result The parameter to place the parsed date into.
' @return A boolean indicating if the date was successfully parsed.
' @remarks This method will not throw an exception if the string cannot be parsed.
'
Public Function TryParse(ByRef s As String, ByRef Result As cDateTime) As Boolean
    On Error GoTo errTrap
    Set Result = Parse(s)
    TryParse = True
errTrap:
End Function

''
' Creates a cDateTime object from a 64-bit currency value.
'
' @param Value The value to initialize the cDateTime with.
' @return A new cDateTime object.
' @remarks The 64-bit value contains both the time in 100-nanosecond
' increments and the DateTimeKind value used to determine if the
' value is a UTC, Local Time, or unspecified time.
' <p>The upper 2-bits represent the DateTimeKind value.</p>
'
Public Function FromBinary(ByVal Value As Currency) As cDateTime
    Set FromBinary = New cDateTime
    FromBinary.InitFromBinary Value
End Function

''
' Creates a new cDateTime object with the specified DateTimeKind.
'
' @param Value The cDateTime to convert to a specific kind.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return A new cDateTime with the original date and new kind.
'
Public Function SpecifyKind(ByVal Value As cDateTime, ByVal Kind As DateTimeKind) As cDateTime
    Set SpecifyKind = New cDateTime
    SpecifyKind.InitFromMilliseconds Value.TotalMilliseconds, Kind
End Function

Public Function Equals(ByRef T1 As Variant, ByRef T2 As Variant) As Boolean
    Dim MS1 As Currency
    Dim MS2 As Currency
    
    If Not TryGetMilliseconds(T1, MS1) Then
        Exit Function
    End If
    
    If Not TryGetMilliseconds(T2, MS2) Then
        Exit Function
    End If
    
    Equals = MS1 = MS2
End Function

Private Function TryGetMilliseconds(ByRef Value As Variant, ByRef Result As Currency) As Boolean
    Select Case VarType(Value)
        Case vbDate
            Result = GetMilliseconds(Value)
            TryGetMilliseconds = True
        
        Case vbObject
            If Value Is Nothing Then
                Exit Function
            End If
            
            If TypeOf Value Is cDateTime Then
                Result = GetMillisecondsFromDateTime(Value)
                TryGetMilliseconds = True
            End If
    End Select
End Function

Private Function GetMillisecondsFromDateTime(ByVal d As cDateTime) As Currency
    GetMillisecondsFromDateTime = d.TotalMilliseconds
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetMilliseconds(ByVal d As Date) As Currency
    Dim Days As Currency
    
    If d < 0# Then
        Days = Fix(d * MILLISECONDS_PER_DAY - 0.5)
        Days = Days - Modulus(Days, MILLISECONDS_PER_DAY) * 2
    Else
        Days = Fix(d * MILLISECONDS_PER_DAY + 0.5)
    End If
    GetMilliseconds = Days + MILLISECONDS_TO_18991231
End Function

Friend Function AscDateTime(ByVal d As Date) As cDateTime
    Set AscDateTime = mDateAscDateTime
    AscDateTime.InitFromDate d, UnspecifiedKind
End Function

''
' This is used to get a valid cDateTime value.
'
' @param d The value to conver to cDateTime
' @return A cDateTime value.
' @remarks If d is already a cDateTime object, then it is returned.
' If d is a Date value, then a new cDateTime object is created.
'
Friend Function GetcDateTime(ByRef d As Variant) As cDateTime
    Select Case VarType(d)
        Case vbObject
            If TypeOf d Is cDateTime Then
                Set GetcDateTime = d
            Else
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired))
            End If
        Case vbDate
            Set GetcDateTime = FromOADate(d)
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired))
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDaysInMonth = Cor.NewLongs(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    Set mDateAscDateTime = New cDateTime
End Sub

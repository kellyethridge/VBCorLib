VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorDateTimeStatic"
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
' Module: CorDateTimeStatic
'

''
' Represents methods to create <a href="CorDateTime.html">CorDateTime</a> objects from various data sources.
'
' @remarks These methods are considered static. The <a href="CorDateTimeStatic.html">CorDateTimeStatic</a> cannot be directly instantiated.
' to access the methods use the <b>CorDateTime</b> qualifier (e.g. CorDateTime.MaxValue.)
' <pre>
' Dim dt As CorDateTime
' Set dt = CorDateTime.FromOADate(#1/1/2001 8:30:00 AM#)
'
' ' A quicker way to create a <a href="CorDateTime.html">CorDateTime</a> object is to use the <a href="Constructors.NewDate.html">NewDate</a> <a href="Constructors.html">constructor</a>.
' Set dt = NewDate(#1/1/2001 8:30:00 AM#)
' </pre>
' This will create a new instance of <a href="CorDateTime.html">CorDateTime</a> with the supplied date value.
' @see CorDateTime
'
'@Folder("CorLib.System")
Option Explicit

Private Const MinMilliseconds As Currency = 0
Private Const MaxMilliseconds As Currency = 315537897599999.9999@


''
' This enumeration represents the days of the week as used by <a href="CorDateTime.html">CorDateTime</a> and <a href="Calendar.html">Calendars</a>.
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
' Used to specify the type of <a href="CorDateTime.html">CorDateTime</a> object.
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
Private mMinValue           As CorDateTime
Private mMaxValue           As CorDateTime
Private mCachedDateTime     As New CorDateTime


''
' Returns a boolean indicating if the specific Gregorian year is a leap year.
'
' @param Year The year to test for leap year.
' @return Indication of leap year.
' @remarks Only the Gregorian calendar is supported through the
' <a href="CorDateTime.html">CorDateTime</a> functions. To determine leap years for other calendar
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
' Returns a <a href="CorDateTime.html">CorDateTime</a> object from a VB Date value.
'
' @param d The standard VB Date value.
' @param Kind Specifies whether the <a href="CorDateTime.html">CorDateTime</a> is a UTC, Local, or unspecified time type.
' @return The <a href="CorDateTime.html">CorDateTime</a> representation of the date.
'
Public Function FromOADate(ByVal d As Date, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As CorDateTime
    Set FromOADate = New CorDateTime
    FromOADate.InitFromDate d, Kind
End Function

''
' Creates a <a href="CorDateTime.html">CorDateTime</a> object from an operating system file timestamp.
'
' @param time The file time in UTC time.
' @return The <a href="CorDateTime.html">CorDateTime</a> of the file time converted to local time from UTC time.
' @remarks Time is a 64-bit integer representing a Windows file timestamp. Using
' a Currency datatype means the timestamp is represented as milliseconds with the
' fraction representing ticks of a millisecond. No conversion if performed. The
' 64-bit values is used as is.
' <p>The timestamp is the number of milliseconds (including the fractional tick portion)
' elasped since January 1, 1601 A.D. 12:00 A.M. coordinated universal time (UTC).</p>
' <p>A tick is 100-nanoseconds. There are 10000 ticks per millisecond, equivalent to
' the number of fractional values in a Currency datatype.</p>
'
Public Function FromFileTime(ByVal Time As Currency) As CorDateTime
    Set FromFileTime = FromFileTimeUtc(Time).ToLocalTime
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> object from a FileTime value.
'
' @param time The file time in UTC time.
' @return The <a href="CorDateTime.html">CorDateTime</a> of the file time in UTC time.
'
Public Function FromFileTimeUtc(ByVal Time As Currency) As CorDateTime
    If Time < 0 Then _
        Error.ArgumentOutOfRange "Time", ArgumentOutOfRange_NeedNonNegNum
    
    Set FromFileTimeUtc = New CorDateTime
    FromFileTimeUtc.InitFromFileTime Time, UtcKind
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> object from milliseconds.
'
' @param ms The number of milliseconds the time is equal to.
' @param Kind Specifies whether the <a href="CorDateTime.html">CorDateTime</a> is a UTC, Local, or unspecified time type.
' @return The <a href="CorDateTime.html">CorDateTime</a> object converted from milliseconds.
' @remarks Since a Currency can be passed in, the fractional portion can be used to represent the number ticks (1/10000 ms) in the date.
'
Public Function FromMilliseconds(ByVal ms As Currency, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As CorDateTime
    Set FromMilliseconds = New CorDateTime
    FromMilliseconds.InitFromMilliseconds ms, Kind
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that represents the current time on the local machine.
'
' @return The <a href="CorDateTime.html">CorDateTime</a> object representing the current time on the local machine.
' @remarks This method uses Windows GetLocalTime which has a resolution down to the
' milliseconds. If the milliseconds is not wanted, then either the Millisecond
' must be subtracted from the instance, or use <a href="CorDateTimeStatic.FromOADate.html">CorDateTime.FromOADate</a> passing the VB Now value.
'
Public Property Get Now() As CorDateTime
    Dim LocalTime As SYSTEMTIME
    
    GetLocalTime LocalTime
    
    With LocalTime
        Set Now = Cor.NewDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds, Kind:=LocalKind)
    End With
End Property

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that represents the current time with the UTC offset.
'
' @return The <a href="CorDateTime.html">CorDateTime</a> object representing the current time with the UTC offset.
' @remarks This method uses Windows GetSystemTime which has a resolution down to the milliseconds. There is no direct VB equivalent.
'
Public Property Get UtcNow() As CorDateTime
    Dim UtcTime As SYSTEMTIME
    
    GetSystemTime UtcTime
    
    With UtcTime
        Set UtcNow = Cor.NewDateTime(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond, .wMilliseconds, Kind:=UtcKind)
    End With
End Property

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that represents only the current date. The time
' portion is not present.
'
' @return The <a href="CorDateTime.html">CorDateTime</a> representing the date.
' @remarks The time is still there, just is only 12:00AM.
'
Public Property Get Today() As CorDateTime
    Dim LocalTime As SYSTEMTIME
    
    GetLocalTime LocalTime
    
    With LocalTime
        Set Today = Cor.NewDateTime(.wYear, .wMonth, .wDay, Kind:=LocalKind)
    End With
End Property

''
' Returns the minimum <a href="CorDateTime.html">CorDateTime</a> representation.
'
' @return The minimum <a href="CorDateTime.html">CorDateTime</a>.
' @remarks The minimum <a href="CorDateTime.html">CorDateTime</a> is 1/1/0001 12:00:00 AM.
'
Public Property Get MinValue() As CorDateTime
    Set MinValue = mMinValue
End Property

''
' Returns the maximum <a href="CorDateTime.html">CorDateTime</a> representation.
'
' @return The maximum <a href="CorDateTime.html">CorDateTime</a>.
' @remarks The maximum <a href="CorDateTime.html">CorDateTime</a> is "12/31/9999 11:59:59 PM".
'
Public Property Get MaxValue() As CorDateTime
    Set MaxValue = mMaxValue
End Property

''
' Parses a string representation of a date into a <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param s The date string to be parsed.
' @return A <a href="CorDateTime.html">CorDateTime</a> object created from the date string.
' @remarks The parse method simply uses CDate to parse the string.
'
Public Function Parse(ByRef s As String) As CorDateTime
    Set Parse = FromOADate(CDate(s))
End Function

''
' Attempts to parse a string representation of a date into a <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param s The date string to be parsed.
' @param result The parameter to place the parsed date into.
' @return A boolean indicating if the date was successfully parsed.
' @remarks This method will not throw an exception if the string cannot be parsed.
'
Public Function TryParse(ByRef s As String, ByRef Result As CorDateTime) As Boolean
    On Error GoTo Catch
    Set Result = Parse(s)
    TryParse = True
Catch:
End Function

''
' Creates a <a href="CorDateTime.html">CorDateTime</a> object from a 64-bit currency value.
'
' @param Value The value to initialize the <a href="CorDateTime.html">CorDateTime</a> with.
' @return A new <a href="CorDateTime.html">CorDateTime</a> object.
' @remarks The 64-bit value contains both the time in 100-nanosecond increments and the <a href="CorDateTimeStatic.DateTimeKind.html">DateTimeKind</a> value used to determine if the value is a UTC, Local Time, or unspecified time.
' <p>The upper 2-bits represent the <a href="CorDateTimeStatic.DateTimeKind.html">DateTimeKind</a> value.</p>
'
Public Function FromBinary(ByVal Value As Currency) As CorDateTime
    Set FromBinary = New CorDateTime
    FromBinary.InitFromBinary Value
End Function

''
' Creates a new <a href="CorDateTime.html">CorDateTime</a> object with the specified <a href="CorDateTimeStatic.DateTimeKind.html">DateTimeKind</a>.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> to convert to a specific kind.
' @param Kind Specifies whether the <a href="CorDateTime.html">CorDateTime</a> is a UTC, Local, or unspecified time type.
' @return A new <a href="CorDateTime.html">CorDateTime</a> with the original date and new kind.
'
Public Function SpecifyKind(ByVal Value As CorDateTime, ByVal Kind As DateTimeKind) As CorDateTime
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"

    If Value.Kind <> Kind Then
        Set SpecifyKind = New CorDateTime
        SpecifyKind.InitFromMilliseconds Value.TotalMilliseconds, Kind
    Else
        Set SpecifyKind = Value
    End If
End Function

''
' Returns a value indicating whether two Date or <a href="CorDateTime.html">CorDateTime</a> instances have the same date and time value.
'
' @param t1 The first Date or <a href="CorDateTime.html">CorDateTime</a> to compare.
' @param t2 The second Date or <a href="CorDateTime.html">CorDateTime</a> to compare.
' @return <b>true</b> if their values are equal; otherwise, <b>false</b>
' @remarks t1 and t2 are equal if their <a href="CorDateTime.Get.TotalMilliseconds.html">TotalMilliseconds</a> property values are equal. Their <a href="CorDateTime.Get.Kind.html">Kind</a> property values are not considered in the test for equality.
'
Public Function Equals(ByRef t1 As Variant, ByRef t2 As Variant) As Boolean
'    Dim d1 As CorDateTime
'    Dim d2 As CorDateTime
'
'    Set d1 = Convert.ToDateTime(t1)
'    Set d2 = Convert.ToDateTime(t2)
'
'    Equals = d1.TotalMilliseconds = d2.TotalMilliseconds
    Dim ms1 As Currency
    Dim ms2 As Currency
    
    On Error GoTo Catch
    ms1 = GetTotalMilliseconds(t1)
    ms2 = GetTotalMilliseconds(t2)
    Equals = ms1 = ms2
    
Catch:
End Function

Friend Function AddYears(ByRef Time As Variant, ByVal Years As Long) As CorDateTime
    Set AddYears = CachedDateTime(Time).AddYears(Years)
End Function

Friend Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As CorDateTime
    Set AddMonths = CachedDateTime(Time).AddMonths(Months)
End Function

Friend Function AddDays(ByRef Time As Variant, ByVal Days As Double) As CorDateTime
    Set AddDays = CachedDateTime(Time).AddDays(Days)
End Function

Friend Function AddHours(ByRef Time As Variant, ByVal Hours As Double) As CorDateTime
    Set AddHours = CachedDateTime(Time).AddHours(Hours)
End Function

Friend Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As CorDateTime
    Set AddMilliseconds = CachedDateTime(Time).AddMilliseconds(Milliseconds)
End Function

Friend Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Double) As CorDateTime
    Set AddMinutes = CachedDateTime(Time).AddMinutes(Minutes)
End Function

Friend Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Double) As CorDateTime
    Set AddSeconds = CachedDateTime(Time).AddSeconds(Seconds)
End Function

Friend Function GetHour(ByRef Time As Variant) As Long
    GetHour = CachedDateTime(Time).Hour
End Function

Friend Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = CachedDateTime(Time).Minute
End Function

Friend Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = CachedDateTime(Time).Second
End Function

Friend Function GetMilliseconds(ByRef Time As Variant) As Long
    GetMilliseconds = CachedDateTime(Time).Millisecond
End Function

Friend Function GetTotalMilliseconds(ByRef Time As Variant) As Currency
    GetTotalMilliseconds = CachedDateTime(Time).TotalMilliseconds
End Function

Friend Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = CachedDateTime(Time).DayOfWeek
End Function

Friend Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDayOfMonth = CachedDateTime(Time).Day
End Function

Friend Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDayOfYear = CachedDateTime(Time).DayOfYear
End Function

Friend Function GetMonth(ByRef Time As Variant) As Long
    GetMonth = CachedDateTime(Time).Month
End Function

Friend Function GetYear(ByRef Time As Variant) As Long
    GetYear = CachedDateTime(Time).Year
End Function

Friend Function GetTimeOfDay(ByRef Time As Variant) As TimeSpan
    Set GetTimeOfDay = CachedDateTime(Time).TimeOfDay
End Function

Friend Function GetOADate(ByRef Time As Variant) As Date
    GetOADate = CachedDateTime(Time).ToOADate
End Function

Friend Sub Validate(ByRef Time As Variant)
    CachedDateTime Time
End Sub

Friend Function GetTotalDays(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    If Year = 0 And Month = 0 And Day = 0 Then
        Exit Function
    End If
    
    If Year < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    If Month < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    If Day < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    
    Dim Days As Long
    If CorDateTime.IsLeapYear(Year) Then
        Days = DaysToMonthLeapYear(Month - 1)
    Else
        Days = DaysToMonth(Month - 1)
    End If
    
    Year = Year - 1
    GetTotalDays = (Day - 1) + Days + 365 * Year + (Year \ 4) - ((Year \ 100) - (Year \ 400))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mDaysInMonth = Cor.NewLongs(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    Set mMinValue = FromMilliseconds(MinMilliseconds)
    Set mMaxValue = FromMilliseconds(MaxMilliseconds)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CachedDateTime(ByRef Time As Variant) As CorDateTime
    Dim Milliseconds As Currency

    Select Case VarType(Time)
        Case vbObject
            If Time Is Nothing Then
                Milliseconds = 0
            ElseIf TypeOf Time Is CorDateTime Then
                Dim dt As CorDateTime
                Set dt = Time
                Milliseconds = dt.TotalMilliseconds
            Else
                Error.Argument Arg_MustBeDateTime
            End If
        Case vbDate
            Milliseconds = DateToMilliseconds(Time)
        Case vbCurrency
            Milliseconds = Time
        Case Else
            Error.Argument Arg_MustBeDateTime
    End Select

    mCachedDateTime.InitFromMilliseconds Milliseconds, UnspecifiedKind

    Set CachedDateTime = mCachedDateTime
End Function

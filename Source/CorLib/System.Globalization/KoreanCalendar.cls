VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KoreanCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2018 Kelly Ethridge
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
' Module: KoreanCalendar
'

''
' Provides functions to manipulate Korean dates.
'
' @see KoreanStatic
' @see Calendar
'
Option Explicit
Implements IObject
Implements Calendar

Private Const PropTwoDigitYearMax   As String = "TwoDigitYearMax"
Private Const MinKoreanYear         As Long = 2334
Private Const MaxKoreanYear         As Long = 12332
Private Const YEAR_OFFSET           As Long = 2333

Private mTwoDigitYearMax    As Long
Private mIsReadOnly         As Boolean


''
' Creates a new object that is a copy of the current <a href="KoreanCalendar.html">KoreanCalendar</a> object.
'
' @return A copy of the <a href="KoreanCalendar.html">KoreanCalendar</a> object.
' @remarks The clone will always be a non-readonly version.
'
Public Function Clone() As KoreanCalendar
    Set Clone = New KoreanCalendar
    Clone.InitClone Me
End Function

''
' Gets a value that indicates whether the current calendar is solar-based, lunar-based, or a combination of both.
'
' @return Always returns <a href="Calendar.CalendarAlgorithmType.html">CalendarAlgorithmType.SolarCalendar</a>.
'
Public Property Get AlgorithmType() As CalendarAlgorithmType
    AlgorithmType = SolarCalendar
End Property

''
' Gets a value indicating whether this <a href="KoreanCalendar.html">KoreanCalendar</a> object is read-only.
'
' @return <b>True</b> if this <a href="KoreanCalendar.html">KoreanCalendar</a> object is read-only; otherwise, <b>False</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Gets a list of eras in the <a href="KoreanCalendar.html">KoreanCalendar</a>.
'
' @return An array of integers that represents the eras in the <a href="KoreanCalendar.html">KoreanCalendar</a>.
' @remarks The <a href="KoreanCalendar.html">KoreanCalendar</a> class recognizes only the current era. This property always returns an array with only one element.
'
Public Property Get Eras() As Long()
    Eras = Cor.NewLongs(1)
End Property

''
' Gets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @return The last year of a 100-year range that can be represented by a 2-digit year.
' @remarks <p>This property allows a 2-digit year to be properly translated to a 4-digit year. For example, in the Gregorian calendar, if this property is set to 2029, the 100-year range is from 1930 to 2029. Therefore, a 2-digit value of 30 is interpreted as 1930, while a 2-digit value of 29 is interpreted as 2029.</p>
' <p>The initial value of this property is derived from the settings in the regional and language options portion of Control Panel. However, that information can change during the life of the application. The <a href="KoreanCalendar.html">KoreanCalendar</a> class does not detect changes in the system settings automatically.</p>
' <p>The application should set this value to 99 to indicate that 2-digit years are to be taken literally. For example, if this property is set to 99, the 100-year range is from 0 (not a valid value for most calendars) to 99. Therefore, a 2-digit value of 30 is interpreted as 30.</p>
'
Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

''
' Sets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @param Value The last year of a 100-year range that can be represented by a 2-digit year.
' @remarks
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>
'           The value specified in a set operation is less than 99.<br/>
'           -or-<br/>
'           The value specified in a set operation is greater than <b>MaxSupportedDateTime.Year</b>.
'       </td>'
'   </tr>
'   <tr>
'       <td><a href="InvalidOperationException.html">InvalidOperationException</a></td>
'       <td>In a set operation, the current instance is read-only.</td>
' </table>
' <p>This property allows a 2-digit year to be properly translated to a 4-digit year. For example, in the Gregorian calendar, if this property is set to 2029, the 100-year range is from 1930 to 2029. Therefore, a 2-digit value of 30 is interpreted as 1930, while a 2-digit value of 29 is interpreted as 2029.</p>
' <p>The initial value of this property is derived from the settings in the regional and language options portion of Control Panel. However, that information can change during the life of the application. The <a href="KoreanCalendar.html">KoreanCalendar</a> class does not detect changes in the system settings automatically.</p>
' <p>The application should set this value to 99 to indicate that 2-digit years are to be taken literally. For example, if this property is set to 99, the 100-year range is from 0 (not a valid value for most calendars) to 99. Therefore, a 2-digit value of 30 is interpreted as 30.</p>
'
Public Property Let TwoDigitYearMax(ByVal Value As Long)
    VerifyWritable
    
    If Value < 99 Or Value > MaxKoreanYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Message:=Environment.GetResourceString(ArgumentOutOfRange_Range, 99, MaxKoreanYear))
    
    mTwoDigitYearMax = Value
End Property

''
' Gets the earliest date and time supported by the <a href="KoreanCalendar.html">KoreanCalendar</a> class.
'
' @return The earliest date and time supported by the <a href="KoreanCalendar.html">KoreanCalendar</a> class, which is equivalent to the first moment of January 1, 0001 C.E. in the Gregorian calendar.
'
Public Property Get MinSupportedDateTime() As CorDateTime
    Set MinSupportedDateTime = CorDateTime.MinValue
End Property

''
' Gets the latest date and time supported by the <a href="KoreanCalendar.html">KoreanCalendar</a> class.
'
' @return The latest date and time supported by the <a href="KoreanCalendar.html">KoreanCalendar</a> class, which is equivalent to the last moment of December 31, 9999 C.E. in the Gregorian calendar.
'
Public Property Get MaxSupportedDateTime() As CorDateTime
    Set MaxSupportedDateTime = CorDateTime.MaxValue
End Property

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that is the specified number of days away from the specified <a href="CorDateTime.html">CorDateTime</a>.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to add days to.
' @param Days The number of days to add.
' @return The <a href="CorDateTime.html">CorDateTime</a> that results from adding the specified number of days to the specified <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As CorDateTime
    Set AddDays = Statics.Calendar.AddDays(Me, Time, Days)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of hours from the original Time.
'
' @param Time The time to add the specified number of hours to.
' @param Hours The number of hours to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of hours added.
'
Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As CorDateTime
    Set AddHours = Statics.Calendar.AddHours(Me, Time, Hours)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Milliseconds from the original Time.
'
' @param Time The time to add the specified number of Milliseconds to.
' @param Milliseconds The number of Milliseconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Milliseconds added.
'
Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As CorDateTime
    Set AddMilliseconds = Statics.Calendar.AddMilliseconds(Me, Time, Milliseconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Minutes from the original Time.
'
' @param Time The time to add the specified number of Minutes to.
' @param Minutes The number of Minutes to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Minutes added.
'
Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As CorDateTime
    Set AddMinutes = Statics.Calendar.AddMinutes(Me, Time, Minutes)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Months from the original Time.
'
' @param Time The time to add the specified number of Months to.
' @param Months The number of Months to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Months added.
'
Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As CorDateTime
    Set AddMonths = Statics.Calendar.AddMonths(Me, Time, Months)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Seconds from the original Time.
'
' @param Time The time to add the specified number of Seconds to.
' @param Seconds The number of Seconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Seconds added.
'
Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As CorDateTime
    Set AddSeconds = Statics.Calendar.AddSeconds(Me, Time, Seconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Weeks from the original Time.
'
' @param Time The time to add the specified number of Weeks to.
' @param Weeks The number of Weeks to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Weeks added.
'
Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As CorDateTime
    Set AddWeeks = AddDays(Time, Weeks * DaysPerWeek)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Years from the original Time.
'
' @param Time The time to add the specified number of Years to.
' @param Years The number of Years to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Years added.
'
Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As CorDateTime
    Set AddYears = Statics.Calendar.AddYears(Me, Time, Years)
End Function

''
' Returns the day of the months for the specified date.
'
' @param Time The date inwhich to retrieve the day of the month.
' @return An integer from 1 to 31 that represents the day of the month in the specified Date or <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDayOfMonth = CorDateTime.GetDayOfMonth(Time)
End Function

''
' Returns the <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> enum of the specified date.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to read the day of the week from.
' @return A <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> enum representing the day of the week.
' @remarks The <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> is Sunday and is 0-based. The Visual Basic
' enum vbDayOfWeek is 1-based.
'
Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = CorDateTime.GetDayOfWeek(Time)
End Function

''
' Returns the day of the year for the specified date.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to read the day of the year from.
' @return An integer from 1 to 366 that represents the day of the year in the specified Date or <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDayOfYear = CorDateTime.GetDayOfYear(Time)
End Function

''
' Returns the number of days in the specified month.
'
' @param Year The year inwhich the month exists. This is used by calendars that
' have months that change based on the year, such as leap years.
' @param Month The month to retrieve the number of days of.
' @param Era The time period inwhich the year exists.
'
Public Function GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Long
    Year = GetGregorianYear(Year, Era)
    GetDaysInMonth = CorDateTime.DaysInMonth(Year, Month)
End Function

''
' Returns the number of days in the specified year.
'
' @param Year The year to retrieve the number of days of.
' @param Era The era the year exists in.
' @return The number of days in the year.
'
Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Year = GetGregorianYear(Year, Era)
    
    If IsLeapYear(Year) Then
        GetDaysInYear = 366
    Else
        GetDaysInYear = 365
    End If
End Function

''
' Returns the era in the specified <a href="CorDateTime.html">CorDateTime</a>.
'
' @param The <a href="CorDateTime.html">CorDateTime</a> to read.
' @return An integer that represents the era in the specified <a href="CorDateTime.html">CorDateTime</a>.
' @remarks The <a href="KoreanCalendar.html">KoreanCalendar</a> class recognizes only the current era.
'
Public Function GetEra(ByRef Time As Variant) As Long
    CorDateTime.Validate Time
    GetEra = 1
End Function

''
' Returns the hour of the specified time.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to retrieve the hour from.
' @return The hour portion of the time.
'
Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = CorDateTime.GetHour(Time)
End Function

''
' Returns the milliseconds of the specified time.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to retrieve the milliseconds from.
' @return The millisecond portion of the time.
'
Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = CorDateTime.GetMilliseconds(Time)
End Function

''
' Returns the minute of the specified time.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to retrieve the minute from.
' @return The minute portion of the time.
'
Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = CorDateTime.GetMinute(Time)
End Function

''
' Returns the month of the specified time.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to retrieve the month from.
' @return The month portion of the time.
'
Public Function GetMonth(ByRef Time As Variant) As Long
    GetMonth = CorDateTime.GetMonth(Time)
End Function

''
' Returns the number of months in the specified year.
'
' @param Year The year to get the number of months of.
' @param Era The era the year exists in.
' @return The number of months in the year.
' @remarks The Korean calendar has 12 months in a year.
'
Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    GetGregorianYear Year, Era
    GetMonthsInYear = 12
End Function

''
' Returns the second of the specified time.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to retrieve the second from.
' @return The second portion of the time.
'
Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = CorDateTime.GetSecond(Time)
End Function

''
' Returns the week of the year that the specified date belongs to.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to find the week it belongs.
' @param Rule Rule to decide what constitutes the first week of the year.
' @param FirstDayOfWeek Defines which week day is the start of a new week.
' @return The week that the date belongs to.
'
Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = Statics.Calendar.GetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
End Function

''
' Returns the year of the specified time.
'
' @param Time The time to retrieve the year from.
' @return The year portion of the time.
'
Public Function GetYear(ByRef Time As Variant) As Long
    GetYear = CorDateTime.GetYear(Time) + YEAR_OFFSET
End Function

''
' Returns a boolean indicating if the specific day is a leap day.
'
' @param Year The year the day is in.
' @param Month The month the day is in.
' @param Day The day to check if is a leap day.
' @param Era The era the year is in.
' @return Boolean indication if the day is a leap day.
' @remarks A leap day only exists in a leap year, such as February 29th in the Korean calendar.
'
Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    VerifyMonth Month
    
    If Day < 1 Or Day > GetDaysInMonth(Year, Month) Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, GetDaysInMonth(Year, Month)))
    
    If Month = 2 And Day = 29 Then
        IsLeapDay = IsLeapYear(Year, Era)
    End If
End Function

''
' Returns a boolean indicating if the specific month is a leap month.
'
' @param Year The year the month is in.
' @param Month The month to check if it is a leap month.
' @param Era The era the year is in.
' @return Boolean indication if the month is a leap month.
' @remarks A leap month only exists in a leap year, such as Adar II in the Hebrew calendar.
' This always returns <b>False</b>.
'
Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    GetGregorianYear Year, Era
    VerifyMonth Month
    
    IsLeapMonth = False
End Function

''
' Returns a boolean indicating if the specific year is a leap year.
'
' @param Year The year to check if is a leap year.
' @param Era The era the year is in.
' @return Boolean indication if the year is a leap year.
' @remarks A leap year in the Korean calendar occurs when the year is
' evenly divisible by 4, but not by 100, unless it is divisible by 400.
'
Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    IsLeapYear = CorDateTime.IsLeapYear(GetGregorianYear(Year, Era))
End Function

''
' Returns a Gregorian <a href="CorDateTime.html">CorDateTime</a> computed using the specific calendar rules.
'
' @param Year The Year in the specific calendar type.
' @param Month The Month in the specific calendar type.
' @param Day The Day in the specific calendar type.
' @param Hour The hour for the new time.
' @param Minute The minute for the new time.
' @param Second The second for the new time.
' @param Millisecond The milliseconds for the new time.
' @param Era The era the year is in.
' @return A Gregorian version of the date specified.
'
Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByRef Era As Variant) As CorDateTime
    Year = GetGregorianYear(Year, Era)
    Set ToDateTime = Cor.NewDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond)
    Statics.Calendar.ValidateDateTime Me, ToDateTime
End Function

''
' Returns a 4-digit year from a 2-digit number.
'
' @param Year The year to be converted to 4-digits.
' @return A 4-digit year in the specific calendar.
' @remarks A normal way of dealing with years is by refering to them as a
' 2-digit value. This is evident in the Gregorian calendar system. The
' year 2005 may be referenced as '05 (pronounced oh-five). When converting
' a 2-digit year to a 4-digit the century is unknown. '05 could be 2005 or 1905.
' This property gives the calendar a rule to follow when converting 2-digit years
' to 4-digit.
' <p>A max year of 2029 will cause a minimum year of 1930. These two years represent
' 100 years which 2-digit years will fall. 00-99 will fall within 1930-2029.</p>
'
Public Function ToFourDigitYear(ByVal Year As Long) As Long
    If (Year > 99 And Year < MinKoreanYear) Or Year > MaxKoreanYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, 0, MaxKoreanYear))
    
    ToFourDigitYear = Statics.Calendar.ToFourDigitYear(Me, Year)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mTwoDigitYearMax = Statics.Calendar.GetCalendarInfoInt32(CAL_KOREA, CAL_ITWODIGITYEARMAX)
End Sub

Friend Sub InitClone(ByVal Other As KoreanCalendar)
    mTwoDigitYearMax = Other.TwoDigitYearMax
End Sub

Friend Sub InitReadOnlyClone(ByVal Other As KoreanCalendar)
    InitClone Other
    mIsReadOnly = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mTwoDigitYearMax = PropBag.ReadProperty(PropTwoDigitYearMax)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PropTwoDigitYearMax, mTwoDigitYearMax
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CloneReadOnly() As KoreanCalendar
    Set CloneReadOnly = New KoreanCalendar
    CloneReadOnly.InitReadOnlyClone Me
End Function

Private Sub VerifyWritable()
    If IsReadOnly Then
        Error.InvalidOperation InvalidOperation_ReadOnly
    End If
End Sub

Private Sub VerifyEra(ByRef Era As Variant)
    Select Case CLngOrDefault(Era, Statics.Calendar.CurrentEra)
        Case KoreanCalendar.KoreanEra, Statics.Calendar.CurrentEra
            ' all good
        Case Else
            Error.Argument Argument_InvalidEraValue, "Era"
    End Select
End Sub

Private Function GetGregorianYear(ByVal Year As Long, ByRef Era As Variant) As Long
    VerifyEra Era
    
    If Year < MinKoreanYear Or Year > MaxKoreanYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, MinKoreanYear, MaxKoreanYear))
    
    GetGregorianYear = Year - YEAR_OFFSET
End Function

Private Sub VerifyMonth(ByVal Month As Long)
    If Month < 1 Or Month > 12 Then
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 12))
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Calendar
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Calendar_AlgorithmType() As CalendarAlgorithmType
    Calendar_AlgorithmType = AlgorithmType
End Property

Private Function Calendar_Clone() As Calendar
    Set Calendar_Clone = Clone
End Function

Private Property Get Calendar_IsReadOnly() As Boolean
    Calendar_IsReadOnly = IsReadOnly
End Property

Private Function Calendar_AddDays(Time As Variant, ByVal Days As Long) As CorDateTime
    Set Calendar_AddDays = AddDays(Time, Days)
End Function

Private Function Calendar_AddHours(Time As Variant, ByVal Hours As Long) As CorDateTime
    Set Calendar_AddHours = AddHours(Time, Hours)
End Function

Private Function Calendar_AddMilliseconds(Time As Variant, ByVal Milliseconds As Double) As CorDateTime
    Set Calendar_AddMilliseconds = AddMilliseconds(Time, Milliseconds)
End Function

Private Function Calendar_AddMinutes(Time As Variant, ByVal Minutes As Long) As CorDateTime
    Set Calendar_AddMinutes = AddMinutes(Time, Minutes)
End Function

Private Function Calendar_AddMonths(Time As Variant, ByVal Months As Long) As CorDateTime
    Set Calendar_AddMonths = AddMonths(Time, Months)
End Function

Private Function Calendar_AddSeconds(Time As Variant, ByVal Seconds As Long) As CorDateTime
    Set Calendar_AddSeconds = AddSeconds(Time, Seconds)
End Function

Private Function Calendar_AddWeeks(Time As Variant, ByVal Weeks As Long) As CorDateTime
    Set Calendar_AddWeeks = AddWeeks(Time, Weeks)
End Function

Private Function Calendar_AddYears(Time As Variant, ByVal Years As Long) As CorDateTime
    Set Calendar_AddYears = AddYears(Time, Years)
End Function

Private Function Calendar_Equals(Value As Variant) As Boolean
    Calendar_Equals = Equals(Value)
End Function

Private Property Get Calendar_Eras() As Long()
    Calendar_Eras = Eras
End Property

Private Function Calendar_GetDayOfMonth(Time As Variant) As Long
    Calendar_GetDayOfMonth = GetDayOfMonth(Time)
End Function

Private Function Calendar_GetDayOfWeek(Time As Variant) As DayOfWeek
    Calendar_GetDayOfWeek = GetDayOfWeek(Time)
End Function

Private Function Calendar_GetDayOfYear(Time As Variant) As Long
    Calendar_GetDayOfYear = GetDayOfYear(Time)
End Function

Private Function Calendar_GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional Era As Variant) As Long
    Calendar_GetDaysInMonth = GetDaysInMonth(Year, Month, Era)
End Function

Private Function Calendar_GetDaysInYear(ByVal Year As Long, Optional Era As Variant) As Long
    Calendar_GetDaysInYear = GetDaysInYear(Year, Era)
End Function

Private Function Calendar_GetEra(Time As Variant) As Long
    Calendar_GetEra = GetEra(Time)
End Function

Private Function Calendar_GetHashCode() As Long
    Calendar_GetHashCode = GetHashCode
End Function

Private Function Calendar_GetHour(Time As Variant) As Long
    Calendar_GetHour = GetHour(Time)
End Function

Private Function Calendar_GetMilliseconds(Time As Variant) As Double
    Calendar_GetMilliseconds = GetMilliseconds(Time)
End Function

Private Function Calendar_GetMinute(Time As Variant) As Long
    Calendar_GetMinute = GetMinute(Time)
End Function

Private Function Calendar_GetMonth(Time As Variant) As Long
    Calendar_GetMonth = GetMonth(Time)
End Function

Private Function Calendar_GetMonthsInYear(ByVal Year As Long, Optional Era As Variant) As Long
    Calendar_GetMonthsInYear = GetMonthsInYear(Year, Era)
End Function

Private Function Calendar_GetSecond(Time As Variant) As Long
    Calendar_GetSecond = GetSecond(Time)
End Function

Private Function Calendar_GetWeekOfYear(Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    Calendar_GetWeekOfYear = GetWeekOfYear(Time, Rule, FirstDayOfWeek)
End Function

Private Function Calendar_GetYear(Time As Variant) As Long
    Calendar_GetYear = GetYear(Time)
End Function

Private Function Calendar_IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapDay = IsLeapDay(Year, Month, Day, Era)
End Function

Private Function Calendar_IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapMonth = IsLeapMonth(Year, Month, Era)
End Function

Private Function Calendar_IsLeapYear(ByVal Year As Long, Optional Era As Variant) As Boolean
    Calendar_IsLeapYear = IsLeapYear(Year, Era)
End Function

Private Property Get Calendar_MaxSupportedDateTime() As CorDateTime
    Set Calendar_MaxSupportedDateTime = MaxSupportedDateTime
End Property

Private Property Get Calendar_MinSupportedDateTime() As CorDateTime
    Set Calendar_MinSupportedDateTime = MinSupportedDateTime
End Property

Private Function Calendar_ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional Era As Variant) As CorDateTime
    Set Calendar_ToDateTime = ToDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond, Era)
End Function

Private Function Calendar_ToFourDigitYear(ByVal Year As Long) As Long
    Calendar_ToFourDigitYear = ToFourDigitYear(Year)
End Function

Private Function Calendar_ToString() As String
    Calendar_ToString = ToString
End Function

Private Property Let Calendar_TwoDigitYearMax(ByVal Value As Long)
    TwoDigitYearMax = Value
End Property

Private Property Get Calendar_TwoDigitYearMax() As Long
    Calendar_TwoDigitYearMax = TwoDigitYearMax
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function



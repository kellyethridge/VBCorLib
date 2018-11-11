VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HebrewCalendar"
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
' Module: HebrewCalendar
'

''
' Provides methods for manipulating Hebrew (Jewish) dates.
'
' @see HebrewCalendarStatic
' @see Calendar
'
Option Explicit
Implements IObject
Implements Calendar
Implements ICloneable

Private Const PROP_TWODIGITYEARMAX As String = "TwoDigitYearMax"

' the Hebrew calendar epoch - that is, Tishri 1, 1 AM -
' corresponds to Monday, October 7, 3761 BC in the Julian calendar, which
' is September 7, 3761 BC in the Gregorian calendar, which is RD -1,373,427.
Private Const EPOCH As Long = -1373427

Private Const SecondsPerMinute      As Long = 60
Private Const MinutesPerHour        As Long = 60
Private Const HoursPerDay           As Long = 24
Private Const MillisecondsPerSecond As Currency = 1000@
Private Const MillisecondsPerMinute As Currency = MillisecondsPerSecond * SecondsPerMinute
Private Const MillisecondsPerHour   As Currency = MillisecondsPerMinute * MinutesPerHour
Private Const MilliSecondsPerDay    As Currency = MillisecondsPerHour * HoursPerDay
Private Const PartsPerHour          As Long = 1080
Private Const MinHebrewYear         As Long = 5343
Private Const MaxHebrewYear         As Long = 5999

Private mIsReadOnly         As Boolean
Private mTwoDigitYearMax    As Long



''
' Gets a value that indicates whether the current calendar is solar-based, lunar-based, or a combination of both.
'
' @return Always returns <a href="Calendar.CalendarAlgorithmType.html">CalendarAlgorithmType.LunisolarCalendar</a>.
'
Public Property Get AlgorithmType() As CalendarAlgorithmType
    AlgorithmType = LunisolarCalendar
End Property

''
' Gets the list of eras in the <a href="HebrewCalendar.html">HebrewCalendar</a>.
'
' @returns An array of integers that represents the eras in the <a href="HebrewCalendar.html">HebrewCalendar</a>.
' @remarks The <a href="HebrewCalendar.html">HebrewCalendar</a> class recognizes only the current era. This property always returns an array with only one element.
'
Public Property Get Eras() As Long()
    Eras = Cor.NewLongs(HebrewCalendar.HebrewEra)
End Property

''
' Gets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @return The last year of a 100-year range that can be represented by a 2-digit year.
' @remarks The <b>TwoDigitYearMax</b> property allows a 2-digit year to be properly translated to a 4-digit year. For example, if this property is set to 5729, the 100-year range is from 5630 to 5729. Therefore, a 2-digit value of 30 is interpreted as 5630, while a 2-digit value of 29 is interpreted as 5729.
' The initial value of this property is derived from the settings in the regional and language options portion of Control Panel. If the initial system setting changes during the life of your application the <a href="HebrewCalendar.html">HebrewCalendar</a> class does not automatically detect the change.
' The special value 99 causes the <a href="HebrewCalendar.ToFourDigitYear.html">ToFourDigitYear</a> method to ignore the system settings and return the specified year unchanged.
Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

''
' Sets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @param Value The last year of a 100-year range.
' @remarks The <b>TwoDigitYearMax</b> property allows a 2-digit year to be properly translated to a 4-digit year. For example, if this property is set to 5729, the 100-year range is from 5630 to 5729. Therefore, a 2-digit value of 30 is interpreted as 5630, while a 2-digit value of 29 is interpreted as 5729.
' The initial value of this property is derived from the settings in the regional and language options portion of Control Panel. If the initial system setting changes during the life of your application the <a href="HebrewCalendar.html">HebrewCalendar</a> class does not automatically detect the change.
' The special value 99 causes the ToFourDigitYear method to ignore the system settings and return the specified year unchanged.
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="InvalidOperationException.html">InvalidOperationException</a></td>
'       <td>This calendar is read-only.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>In a set operation, the Hebrew calendar year value is less than 5343 but is not 99, or the year value is greater than 5999.</td>
'   </tr
' </table>
Public Property Let TwoDigitYearMax(ByVal Value As Long)
    VerifyWritable
    
    If Value <> 99 Then
        ValidateHebrewYear Value
    End If
    
    mTwoDigitYearMax = Value
End Property

''
' Gets a value indicating whether this <a href="HebrewCalendar.html">HebrewCalendar</a> object is read-only.
'
' @return <b>True</b> if this <a href="HebrewCalendar.html">HebrewCalendar</a> object is read-only; otherwise, <b>False</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Creates a new object that is a copy of the current <a href="HebrewCalendar.html">HebrewCalendar</a> object.
'
' @return A copy of the <a href="HebrewCalendar.html">HebrewCalendar</a> object.
' @remarks The clone will always be a non-readonly version.
'
Public Function Clone() As HebrewCalendar
    Set Clone = New HebrewCalendar
    Clone.InitClone mTwoDigitYearMax
End Function

''
' Gets the earliest date and time supported by the <a href="HebrewCalendar.html">HebrewCalendar</a> object.
'
' @return The earliest date and time supported by the <a href="HebrewCalendar.html">HebrewCalendar</a> which is January 1, 1583.
'
Public Property Get MinSupportedDateTime() As CorDateTime
    Set MinSupportedDateTime = HebrewCalendar.MinSupportedDateTime
End Property

''
' Gets the latest date and time supported by the <a href="HebrewCalendar.html">HebrewCalendar</a> object.
'
' @return The latest date and time supported by the <a href="HebrewCalendar.html">HebrewCalendar</a> which is the last moment of September 29, 2239.
'
Public Property Get MaxSupportedDateTime() As CorDateTime
    Set MaxSupportedDateTime = HebrewCalendar.MaxSupportedDateTime
End Property

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that is the specified number of days away from the specified <a href="CorDateTime.html">CorDateTime</a>.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to add days to.
' @param Days The number of days to add.
' @return The <a href="CorDateTime.html">CorDateTime</a> that results from adding the specified number of days to the specified <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As CorDateTime
    Set AddDays = Statics.Calendar.GetDateTime(Time).AddDays(Days)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of hours from the original Time.
'
' @param Time The time to add the specified number of hours to.
' @param Hours The number of hours to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of hours added.
'
Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As CorDateTime
    Set AddHours = Statics.Calendar.GetDateTime(Time).AddHours(Hours)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Milliseconds from the original Time.
'
' @param Time The time to add the specified number of Milliseconds to.
' @param Milliseconds The number of Milliseconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Milliseconds added.
'
Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As CorDateTime
    Set AddMilliseconds = Statics.Calendar.GetDateTime(Time).AddMilliseconds(Milliseconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Minutes from the original Time.
'
' @param Time The time to add the specified number of Minutes to.
' @param Minutes The number of Minutes to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Minutes added.
'
Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As CorDateTime
    Set AddMinutes = Statics.Calendar.GetDateTime(Time).AddMinutes(Minutes)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Months from the original Time.
'
' @param Time The time to add the specified number of Months to.
' @param Months The number of Months to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Months added.
'
Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As CorDateTime
    Dim Year    As Long
    Dim Month   As Long
    Dim Day     As Long
    Dim TOD     As Currency
    Dim ms      As Currency
    
    If Months = 0 Then
        Set AddMonths = CorDateTime.GetDateTime(Time)
        Exit Function
    End If
    
    ms = Statics.Calendar.GetDateTime(Time).TotalMilliseconds
    TOD = Modulus(ms, MilliSecondsPerDay)
    GetDateParts ms, Complete, Year, Month, Day
    Month = ToDotNetMonth(Year, Month)
    
    If Months < 0 Then
        Do While Months < 0
            If Months + Month > 0 Then
                Month = Month + Months
                Exit Do
            Else
                Months = Months + Month
                Year = Year - 1
                Month = GetMonthsInYear(Year)
            End If
        Loop
    Else
        Dim i As Long
        Do While Months > 0
            i = GetMonthsInYear(Year)
            If Months + Month <= i Then
                Month = Month + Months
                Exit Do
            Else
                Months = Months - (i - Month + 1)
                Month = 1
                Year = Year + 1
            End If
        Loop
    End If
    
    Dim Max As Long
    Max = GetDaysInMonth(Year, Month)
    
    If Day > Max Then
        Day = Max
    End If
    
    Set AddMonths = ToDateTime(Year, Month, Day, 0, 0, 0, 0).AddMilliseconds(TOD)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Seconds from the original Time.
'
' @param Time The time to add the specified number of Seconds to.
' @param Seconds The number of Seconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Seconds added.
'
Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As CorDateTime
    Set AddSeconds = Statics.Calendar.GetDateTime(Time).AddSeconds(Seconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Weeks from the original Time.
'
' @param Time The time to add the specified number of Weeks to.
' @param Weeks The number of Weeks to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Weeks added.
'
Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As CorDateTime
    Set AddWeeks = AddDays(Time, Weeks * 7)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Years from the original Time.
'
' @param Time The time to add the specified number of Years to.
' @param Years The number of Years to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Years added.
'
Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As CorDateTime
    Dim Year    As Long
    Dim Month   As Long
    Dim Day     As Long
    Dim ms      As Currency
    
    ms = Statics.Calendar.GetDateTime(Time).TotalMilliseconds
    GetDateParts ms, Complete, Year, Month, Day
    
    Month = ToDotNetMonth(Year, Month)
    Year = Year + Years
    
    Set AddYears = ToDateTime(Year, Month, Day, 0, 0, 0, 0).AddMilliseconds(Modulus(ms, MilliSecondsPerDay))
End Function

''
' Returns the era in the specified Date or <a href="CorDateTime.html">CorDateTime</a>.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to read.
' @return An integer that represents the era in the specified Date or <a href="CorDateTime.html">CorDateTime</a>. The return value is always <a href="HebrewCalendarStatic.Get.HebrewEra.html">HebrewEra</a>.
'
Public Function GetEra(ByRef Time As Variant) As Long
    Statics.Calendar.GetDateTime Time   ' verifies we have a date.
    GetEra = HebrewCalendar.HebrewEra
End Function

''
' Returns the hour of the specified time.
'
' @param Time The time to retrieve the hour from.
' @return The hour portion of the <i>Time</i>.
'
Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = Statics.Calendar.GetDateTime(Time).Hour
End Function

''
' Returns the milliseconds of the specified time.
'
' @param Time The time to retrieve the milliseconds from.
' @return The millisecond portion of the <i>Time</i>.
'
Public Function GetMilliseconds(ByRef Time As Variant) As Long
    GetMilliseconds = Statics.Calendar.GetMilliseconds(Time)
End Function

''
' Returns the second of the specified time.
'
' @param Time The time to retrieve the second from.
' @return The second portion of the <i>Time</i>.
'
Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = Statics.Calendar.GetSecond(Time)
End Function

''
' Returns the minute of the specified time.
'
' @param Time The time to retrieve the minute from.
' @return The minute portion of the <i>Time</i>.
'
Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = Statics.Calendar.GetMinute(Time)
End Function

''
' Returns the month of the specified time.
'
' @param Time The time to retrieve the month from.
' @return The month portion of the <i>Time</i>.
'
Public Function GetMonth(ByRef Time As Variant) As Long
    GetDateParts Statics.Calendar.GetDateTime(Time).TotalMilliseconds, MonthPart, , GetMonth
End Function

''
' Returns the year of the specified time.
'
' @param Time The time to retrieve the year from.
' @return The year portion of the <i>Time</i>.
'
Public Function GetYear(ByRef Time As Variant) As Long
    GetDateParts Statics.Calendar.GetDateTime(Time).TotalMilliseconds, YearPart, GetYear
End Function

''
' Returns the week of the year that the specified date belongs to.
'
' @param Time The date to find the week it belongs.
' @param Rule Rule to decide what constitutes the first week of the year.
' @param FirstDayOfWeek Defines which week day is the start of a new week.
' @return The week that the date belongs to.
'
Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = Statics.Calendar.GetWeekOfYear(Me, Time, Rule, FirstDayOfWeek)
End Function

''
' Returns the number of months in the specified year.
'
' @param Year The year to get the number of months of.
' @param Era The era the year exists in.
' @return The number of months in the year.
'
Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    If IsLeapYear(Year, Era) Then
        GetMonthsInYear = 13
    Else
        GetMonthsInYear = 12
    End If
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
    ValidateEra Era
    ValidateHebrewMonth Month, Year
    GetDaysInMonth = GetDaysInHebrewMonth(Year, ToHebrewMonth(Year, Month))
End Function

''
' Returns the number of days in the specified year.
'
' @param Year The year to retrieve the number of days of.
' @param Era The era the year exists in.
' @return The number of days in the year.
'
Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    ValidateEra Era
    ValidateHebrewYear Year
    GetDaysInYear = GetFixedDaysFromParts(Year + 1, 7, 1) - GetFixedDaysFromParts(Year, 7, 1)
End Function

''
' Returns the day of the months for the specified date.
'
' @param Time The date inwhich to retrieve the day of the month.
' @return An integer from 1 to 30 that represents the day of the month in the specified Date or <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDateParts Statics.Calendar.GetDateTime(Time).TotalMilliseconds, DayPart, , , GetDayOfMonth
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
    GetDayOfWeek = Statics.Calendar.GetDateTime(Time).DayOfWeek
End Function

''
' Returns the day of the year for the specified date.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to read the day of the year from.
' @return An integer from 1 to 385 that represents the day of the year in the specified Date or <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDateParts Statics.Calendar.GetDateTime(Time).TotalMilliseconds, DayOfTheYear, , , , GetDayOfYear
End Function

''
' Returns a boolean indicating if the specific year is a leap year.
'
' @param Year The year to check if is a leap year.
' @param Era The era the year is in.
' @return Boolean indication if the year is a leap year.
'
Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    ValidateEra Era
    ValidateHebrewYear Year
    IsLeapYear = ((235 * Year - 234) Mod 19) > 11
End Function

''
' Returns a boolean indicating if the specific month is a leap month.
'
' @param Year The year the month is in.
' @param Month The month to check if it is a leap month.
' @param Era The era the year is in.
' @return Boolean indication if the month is a leap month.
' @remarks <p>This implementation of the <a href="HebrewCalendar.html">HebrewCalendar</a> class recognizes only the Hebrew years 5343 to 5999 (1583 to 2239 in the Gregorian calendar).</p>
' <p>In every 19-year cycle that ends with a year that is evenly divisible by 19, the 3rd, 6th, 8th, 11th, 14th, 17th, and 19th years are leap years. A common year can have from 353 to 355 days, depending on the placement of Jewish holidays. A leap year can have from 383 to 385 days.</p>
' <p>A leap month is an entire month that occurs only in a leap year. In the Hebrew calendar, Adar Beit is the only leap month.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>year</i>, <i>month</i>, or <i>era</i> is outside the range supported by this calendar.</td>
'   </tr>
' </table>
Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    ValidateEra Era
    ValidateHebrewMonth Month, Year
    
    If Month = 7 Then
        IsLeapMonth = IsLeapYear(Year)
    End If
End Function

''
' Returns a boolean indicating if the specific day is a leap day.
'
' @param Year The year the day is in.
' @param Month The month the day is in.
' @param Day The day to check if is a leap day.
' @param Era The era the year is in.
' @return Boolean indication if the day is a leap day.
' @remarks <p>This implementation of the HebrewCalendar class recognizes only the Hebrew years 5343 to 5999 (1583 to 2239 in the Gregorian calendar).</p>
' <p>In every 19-year cycle that ends with a year that is evenly divisible by 19, the 3rd, 6th, 8th, 11th, 14th, 17th, and 19th years are leap years. A common year can have from 353 to 355 days, depending on the placement of Jewish holidays. A leap year can have from 383 to 385 days.</p>
' <p>A leap day is a day that occurs only in a leap year. In the Hebrew calendar, the last day of Adar Alef and all the days in Adar Beit are considered leap days.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td><i>year</i>, <i>month</i>, or <i>era</i> is outside the range supported by this calendar.</td>
'   </tr>
' </table>

Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    ValidateHebrewDay Year, Month, Day

    If IsLeapYear(Year, Era) Then
        If Month = 6 Then
            IsLeapDay = (Day = 30)
        Else
            IsLeapDay = IsLeapMonth(Year, Month)
        End If
    End If
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
    ValidateEra Era
    ValidateHebrewDay Year, Month, Day
    
    Dim Days As Long
    Days = GetFixedDaysFromParts(Year, ToHebrewMonth(Year, Month), Day) - 1
    Set ToDateTime = CorDateTime.FromMilliseconds(Days * MilliSecondsPerDay + Hour * MillisecondsPerHour + Minute * MillisecondsPerMinute + Second * MillisecondsPerSecond + Millisecond)
End Function

''
' Converts the specified year to a four-digit year by using the <a href="JulianCalendar.TwoDigitYearMax.get.html">TwoDigitYearMax</a> property to determine the appropriate century.
'
' @param Year A two-digit or four-digit integer that represents the year to convert.
' @return A 4-digit year in the specific calendar.
' @remarks <p>The <b>ToFourDigitYear</b> method uses the year parameter, the <a href="HebrewCalendar.Get.TwoDigitYearMax.html">TwoDigitYearMax</a> property, and a year to calculate a 4-digit year. The century is determined by finding the sole occurrence of the year parameter within that 100-year range. For example, if <a href="HebrewCalendar.Set.TwoDigitYearMax.html">TwoDigitYearMax</a> is set to 5729, the 100-year range is from 5630 to 5729. Therefore, a value of 30 is interpreted as 5630, while a value of 29 is interpreted as 5729.</p>
' <p>If the <a href="HebrewCalendar.Get.TwoDigitYearMax.html">TwoDigitYearMax</a> property has the special value 99, the <a href="HebrewCalendar.ToFourDigitYear.html">ToFourDigitYear</a> method ignores the settings in the regional and language options in Control Panel and returns the value of the year parameter unchanged.</p>
' <p><a href="HebrewCalendar.ToFourDigitYear.html">ToFourDigitYear</a> supports either a two-digit year or a four-digit year. Passing a two-digit year value (less than 100) causes the method to convert the value to a four-digit value according to the <a href="HebrewCalendar.Get.TwoDigitYearMax.html">TwoDigitYearMax</a> value representing the appropriate century. If the application supplies a four-digit year value that is within the supported calendar range to <a href="HebrewCalendar.ToFourDigitYear.html">ToFourDigitYear</a>, the method returns the actual input value. If the application supplies a four-digit value that is outside the supported calendar range, or if it supplies a negative value, the method throws an exception.</p>
'
Public Function ToFourDigitYear(ByVal Year As Long) As Long
    If (Year >= 100 And Year < MinHebrewYear) Or (Year > MaxHebrewYear) Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, MinHebrewYear, MaxHebrewYear))
    
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
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub MarkReadOnly()
    mIsReadOnly = True
End Sub

Private Sub VerifyWritable()
    If IsReadOnly Then
        Error.InvalidOperation InvalidOperation_ReadOnly
    End If
End Sub

Private Sub ValidateEra(ByRef Era As Variant)
    Select Case CLngOrDefault(Era, Statics.Calendar.CurrentEra)
        Case HebrewCalendar.HebrewEra, Statics.Calendar.CurrentEra
            ' all good
        Case Else
            Error.Argument Argument_InvalidEraValue, "Era"
    End Select
End Sub

Private Sub ValidateHebrewYear(ByVal Year As Long)
    If Year < MinHebrewYear Or Year > MaxHebrewYear Then
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, MinHebrewYear, MaxHebrewYear))
    End If
End Sub

Private Sub ValidateHebrewMonth(ByVal Month As Long, ByVal Year As Long)
    Dim MonthsInYear As Long
    
    MonthsInYear = GetMonthsInYear(Year)
    
    If Month < 1 Or Month > MonthsInYear Then
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, MonthsInYear))
    End If
End Sub

Private Sub ValidateHebrewDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long)
    Dim DaysInMonth As Long
    
    DaysInMonth = GetDaysInMonth(Year, Month)
    
    If Day < 1 Or Day > DaysInMonth Then
        Throw Cor.NewArgumentOutOfRangeException("Day", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, DaysInMonth))
    End If
End Sub

' this is based on the algorithm found at "http://www.funaba.org/en/calendar.html"
Private Function GetDaysToHebrewYear(ByVal Year As Long) As Long
    Dim Months      As Long
    Dim Parts       As Long
    Dim Hours       As Long
    Dim Days        As Long
    Dim TotalDays   As Long
    Dim R           As Long
    
    ' months before the specified year.
    Months = ((Year * 235 - 234) \ 19)
    
    ' get the day and parts of a day in the remaining month.
    Days = Months \ PartsPerHour
    R = Months - (Days * PartsPerHour)
    Parts = (R * 793) + 204
    
    ' calculate the total days
    Hours = (Months * 12) + 11 + (Days * 793) + (Parts \ PartsPerHour)
    TotalDays = (Months * 29) + (Hours \ 24)
    
    ' cannot start a year on a Sunday, Wednesday or Friday.
    If (TotalDays + 1) * 3 Mod 7 < 3 Then
        TotalDays = TotalDays + 1
    End If
    
    GetDaysToHebrewYear = TotalDays
End Function

Private Function GetYearStartOffset(ByVal Year As Long) As Long
    Dim Year1 As Long
    Dim Year2 As Long
    
    Year1 = GetDaysToHebrewYear(Year)
    Year2 = GetDaysToHebrewYear(Year + 1)
    
    If Year2 - Year1 = 356 Then
        GetYearStartOffset = 2
    Else
        Year = GetDaysToHebrewYear(Year - 1)
        
        If Year1 - Year = 382 Then
            GetYearStartOffset = 1
        End If
    End If
End Function

''
' Calculations are performed using the month of 'Nissan' as the
' first month. This is the month that a Hebrew year actually
' increments by 1. The first month in Dot NET is Tishrei,
' so the calculated month needs to be adjusted to the Dot NET month.
'
' @param Year The year that contains the month to be mapped. The year may be a leap year which would alter the mapping.
' @param Month The month to be mapped from Hebrew to Dot NET alignment.
' @return The Dot NET version of the month.
'
Private Function ToDotNetMonth(ByVal Year As Long, ByVal HebrewMonth As Long) As Long
    If HebrewMonth >= 7 Then
        ToDotNetMonth = HebrewMonth - 6
    Else
        If IsLeapYear(Year) Then
            ToDotNetMonth = HebrewMonth + 7
        Else
            ToDotNetMonth = HebrewMonth + 6
        End If
    End If
End Function

''
' Calculations are performed using the month of 'Nissan' as the
' first month. This is the month that the Hebrew year actually
' increments by 1. When a function is called by a user and the
' month is a parameter, that month is aligned to the Dot NET version
' of the months and needs to be mapped to the Hebrew version to be
' used correctly in calculations.
'
' @param Year The year that contains the month to be mapped. The year may be a leap year which would alter the mapping.
' @param Month The month to be mapped from Dot NET to Hebrew alignment.
' @return The Hebrew version of the month.
'
Private Function ToHebrewMonth(ByVal Year As Long, ByVal Month As Long) As Long
    If Month < 7 Then
        ToHebrewMonth = Month + 6
    Else
        If IsLeapYear(Year) Then
            If Month = 7 Then
                ToHebrewMonth = Month + 6
            Else
                ToHebrewMonth = Month - 7
            End If
        Else
            ToHebrewMonth = Month - 6
        End If
    End If
End Function

Private Function GetDaysInHebrewMonth(ByVal Year As Long, ByVal Month As Long) As Long
    GetDaysInHebrewMonth = 30
    
    Select Case Month
        Case 2, 4, 6, 10, 13
            GetDaysInHebrewMonth = 29
        Case 8
            If GetDaysInYear(Year) Mod 10 <> 5 Then
                GetDaysInHebrewMonth = 29
            End If
        Case 9
            If GetDaysInYear(Year) Mod 10 = 3 Then
                GetDaysInHebrewMonth = 29
            End If
        Case 12
            If Not IsLeapYear(Year) Then
                GetDaysInHebrewMonth = 29
            End If
    End Select
End Function

Private Sub GetDateParts(ByVal ms As Currency, ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim FixedDays As Long
    
    If ms < MinSupportedDateTime.TotalMilliseconds Or ms > MaxSupportedDateTime.TotalMilliseconds Then _
        Throw Cor.NewArgumentOutOfRangeException("Time", Environment.GetResourceString(ArgumentOutOfRange_CalendarRange, MinSupportedDateTime, MaxSupportedDateTime))
    
    FixedDays = Int(ms / MilliSecondsPerDay) + 1
    
    Year = Int((FixedDays - EPOCH) / 365.246822205978)  ' this is an approximation
    
    Do While FixedDays >= GetFixedDaysFromParts(Year, 7, 1)
        Year = Year + 1
    Loop
    
    Year = Year - 1
    
    If DatePart = YearPart Then
        Exit Sub
    End If
    
    If FixedDays < GetFixedDaysFromParts(Year, 1, 1) Then
        Month = 7
    Else
        Month = 1
    End If
    
    Do While FixedDays > GetFixedDaysFromParts(Year, Month, GetDaysInHebrewMonth(Year, Month))
        Month = Month + 1
    Loop
    
    If DatePart = MonthPart Then
        Month = ToDotNetMonth(Year, Month)
        Exit Sub
    End If
    
    Day = FixedDays - GetFixedDaysFromParts(Year, Month, 1) + 1
    
    If DatePart = DayPart Then
        Exit Sub
    End If
    
    DayOfYear = FixedDays - GetFixedDaysFromParts(Year, 7, 1) + 1
End Sub

Private Function GetFixedDays(ByRef Time As Variant) As Long
    Dim ms As Currency
    ms = Statics.Calendar.GetTotalMilliseconds(Time)
    GetFixedDays = Int(ms / MilliSecondsPerDay) + 1
End Function

Private Function GetFixedDaysFromParts(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    Dim i           As Long
    Dim DayCount    As Long
    Dim MaxMonths   As Long
    
    DayCount = EPOCH - 1 + GetDaysToHebrewYear(Year) + GetYearStartOffset(Year)
    
    If Month < 7 Then
        For i = 1 To Month - 1
            DayCount = DayCount + GetDaysInHebrewMonth(Year, i)
        Next
        
        MaxMonths = GetMonthsInYear(Year)
    Else
        MaxMonths = Month - 1
    End If
    
    For i = 7 To MaxMonths
        DayCount = DayCount + GetDaysInHebrewMonth(Year, i)
    Next
    
    GetFixedDaysFromParts = DayCount + Day
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub InitClone(ByVal TwoDigitYearMax As Long)
    mTwoDigitYearMax = TwoDigitYearMax
    mIsReadOnly = False
End Sub

Private Sub Class_InitProperties()
    mTwoDigitYearMax = Statics.Calendar.GetCalendarInfoInt32(CAL_HEBREW, CAL_ITWODIGITYEARMAX)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mTwoDigitYearMax = PropBag.ReadProperty(PROP_TWODIGITYEARMAX)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_TWODIGITYEARMAX, mTwoDigitYearMax
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


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


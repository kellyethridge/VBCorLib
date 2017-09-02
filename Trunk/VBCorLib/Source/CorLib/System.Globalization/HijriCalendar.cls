VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HijriCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: HijriCalendar
'

''
' Provides methods for manipulating Hijri dates.
'
' @remarks This is based on the Tabular version of the Islamic calendar,
' not the pure lunar version.
' @see HijriCalendarStatic
' @see Calendar
'
Option Explicit
Implements IObject
Implements Calendar
Implements ICloneable

Private Const DefaultTwoDigitYearMax As Long = 1451

' Number of Gregorian days to 7/18/622.
' This is 2 days after the start of the Hijri calendar to
' allow for the Hijri Adjustment to still fit in the window.
' https://fisher.osu.edu/~muhanna.1/hijri-intro.html
Private Const EPOCH As Long = 227013

' The Hijri calendar has a 30 year cycle with 11 of those years being leap years.
' The common year is 354 days, leap years are 355.
' (19 * 354) + (11 * 355) = 10631
Private Const DaysPer30Years        As Long = 10631

' This is used similar to adding 0.5 to a value and taking the integer to
' round to the nearest whole integer.
' DaysPer30Years + (((30 * 355) - (30 * 354))/2) = 10646
Private Const YearRoundingBias      As Long = 10646

' This is 4/3/9666 in fixed days, the maximum valid Hijri date.
Private Const MaxFixedDays          As Long = 3530161
Private Const MaxYear               As Long = 9666
Private Const MaxShortMonth         As Long = 4
Private Const MaxShortDay           As Long = 3
Private Const MonthsPerYear         As Long = 12

Private Const SecondsPerMinute      As Long = 60
Private Const MinutesPerHour        As Long = 60
Private Const HoursPerDay           As Long = 24
Private Const DaysPerWeek           As Long = 7

Private Const MillisecondsPerSecond As Currency = 1000@
Private Const MillisecondsPerMinute As Currency = MillisecondsPerSecond * SecondsPerMinute
Private Const MillisecondsPerHour   As Currency = MillisecondsPerMinute * MinutesPerHour
Private Const MilliSecondsPerDay    As Currency = MillisecondsPerHour * HoursPerDay

Private mHijriAdjustment    As Long
Private mTwoDigitYearMax    As Long
Private mMonthDays()        As Long
Private mIsReadOnly         As Boolean


''
' Gets a value that indicates whether the current calendar is solar-based, lunar-based, or a combination of both.
'
' @return Always returns <a href="Calendar.CalendarAlgorithmType.html">CalendarAlgorithmType.LunarCalendar</a>.
'
Public Property Get AlgorithmType() As CalendarAlgorithmType
    AlgorithmType = LunarCalendar
End Property

''
' Creates a new object that is a copy of the current <a href="HijriCalendar.html">HijriCalendar</a> object.
'
' @return A copy of the <a href="HijriCalendar.html">HijriCalendar</a> object.
'
Public Function Clone() As HijriCalendar
    Set Clone = New HijriCalendar
    Clone.InitClone Me
End Function

''
' Gets the number of days in the year that precedes the year that is specified by the <a href="HijriCalendar.Get.MinSupportedDateTime.html">MinSupportedDateTime</a> property.
'
' @return The number of days in the year that precedes the year specified by <a href="HijriCalendar.Get.MinSupportedDateTime.html">MinSupportedDateTime</a>.
' @remarks This property returns a constant value of 354.
'
Public Property Get DaysInYearBeforeMinSupportedYear() As Long
    DaysInYearBeforeMinSupportedYear = 354
End Property

''
' Gets a value indicating whether this <a href="HijriCalendar.html">HijriCalendar</a> object is read-only.
'
' @return <b>True</b> if this <a href="HijriCalendar.html">HijriCalendar</a> object is read-only; otherwise, <b>False</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Gets the latest date and time supported by this <a href="HijriCalendar.html">HijriCalendar</a> object.
'
' @return The latest date and time supported by the <a href="HijriCalendar.html">HijriCalendar</a> type, which is equivalent to the last moment of December 31, 9999 C.E. in the Gregorian calendar.
' @remarks December 31, 9999 C.E. is equivalent to the 3rd day of the 4th month of the year 9666 in the Hijri calendar.
'
Public Property Get MaxSupportedDateTime() As CorDateTime
    Set MaxSupportedDateTime = HijriCalendar.MaxDateTime
End Property

''
' Gets the earliest date and time supported by this <a href="HijriCalendar.html">HijriCalendar</a> object.
'
' @return The earliest date and time supported by the <a href="HijriCalendar.html">HijriCalendar</a> type, which is equivalent to the first moment of July 18, 622 C.E. in the Gregorian calendar.
'
Public Property Get MinSupportedDateTime() As CorDateTime
    Set MinSupportedDateTime = HijriCalendar.MinDateTime
End Property

''
' Gets the number of days to add or subtract from the calendar to accommodate the variances in the start and the end of Ramadan and to accommodate the date difference between countries/regions.
'
Public Property Get HijriAdjustment() As Long
    HijriAdjustment = mHijriAdjustment
End Property

''
' Sets the number of days to add or subtract from the calendar to accommodate the variances in the start and the end of Ramadan and to accommodate the date difference between countries/regions.
'
' @param Value An integer from -2 to 2 that represents the number of days to add or subtract from the calendar.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>The property is being set to an invalid value.</td>
'   </tr
' </table>
Public Property Let HijriAdjustment(ByVal Value As Long)
    VerifyWritable
    If Value < -2 Or Value > 2 Then _
        Throw Cor.NewArgumentOutOfRangeException("HijriAdjustment", Environment.GetResourceString(ArgumentOutOfRange_Range, -2, 2), Value)
    
    mHijriAdjustment = Value
End Property

''
' Gets a list of eras in the <a href="HijriCalendar.html">HijriCalendar</a>.
'
' @return An array of integers that represents the eras in the <a href="HijriCalendar.html">HijriCalendar</a>.
' @remarks The Hijri calendar recognizes one era: A.H. (Latin "Anno Hegirae", which means "the year of the migration," in reference to the migration of Muhammad (PBUH) from Mecca). This property always returns an array with only one element.
'
Public Property Get Eras() As Long()
    Eras = Cor.NewLongs(HijriCalendar.HijriEra)
End Property

''
' Gets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @return The last year of a 100-year range that can be represented by a 2-digit year.
'
Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

''
' Sets the last year of a 100-year range that can be represented by a 2-digit year.
'
' @param Value The last year of a 100-year range.
' @remarks <p>This property allows a 2-digit year to be properly translated to a 4-digit year. For example, if this property is set to 1429, the 100-year range is from 1330 to 1429. Therefore, a 2-digit value of 30 is interpreted as 1330, while a 2-digit value of 29 is interpreted as 1429.</p>
' <p>The application should set this value to 99 to indicate that 2-digit years are to be taken literally. For example, if this property is set to 99, the 100-year range is from 0 (not a valid value for most calendars) to 99. Therefore, a 2-digit value of 30 is interpreted as 30.</p>
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
'       <td>The value in a set operation is less than 100 or greater than 9666.</td>
'   </tr
' </table>
Public Property Let TwoDigitYearMax(ByVal Value As Long)
    VerifyWritable
    If Value < 100 Or Value > MaxYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, 100, MaxYear))
    
    mTwoDigitYearMax = Value
End Property

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> that is the specified number of days away from the specified <a href="CorDateTime.html">CorDateTime</a>.
'
' @param Time The Date or <a href="CorDateTime.html">CorDateTime</a> to add days to.
' @param Days The number of days to add.
' @return The <a href="CorDateTime.html">CorDateTime</a> that results from adding the specified number of days to the specified <a href="CorDateTime.html">CorDateTime</a>.
'
Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As CorDateTime
    Set AddDays = Statics.Calendar.AddDays(Time, Days)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of hours from the original Time.
'
' @param Time The time to add the specified number of hours to.
' @param Hours The number of hours to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html"> with the specified number of hours added.
'
Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As CorDateTime
    Set AddHours = Statics.Calendar.AddHours(Time, Hours)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Milliseconds from the original Time.
'
' @param Time The time to add the specified number of Milliseconds to.
' @param Milliseconds The number of Milliseconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Milliseconds added.
'
Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As CorDateTime
    Set AddMilliseconds = Statics.Calendar.AddMilliseconds(Time, Milliseconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Minutes from the original Time.
'
' @param Time The time to add the specified number of Minutes to.
' @param Minutes The number of Minutes to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Minutes added.
'
Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As CorDateTime
    Set AddMinutes = Statics.Calendar.AddMinutes(Time, Minutes)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Months from the original Time.
'
' @param Time The time to add the specified number of Months to.
' @param Months The number of Months to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Months added.
'
Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As CorDateTime
    Dim Year        As Long
    Dim Month       As Long
    Dim Day         As Long
    Dim MaxDays     As Long
    Dim FixedDays   As Long
    Dim ms          As Currency
    
    If Months < -120000 Or Months > 120000 Then _
        Throw Cor.NewArgumentOutOfRangeException("Months", Environment.GetResourceString(ArgumentOutOfRange_Range, -120000, 120000))
    
    If Months = 0 Then
        Set AddMonths = CorDateTime.GetDateTime(Time)
        Exit Function
    End If
    
    GetDateParts Time, Complete, Year, Month, Day
    Month = Month + Months
    Year = Year + CorMath.DivRem(Month, MonthsPerYear, Month)
    MaxDays = GetDaysInMonth(Year, Month)
    
    If Day > MaxDays Then
        Day = MaxDays
    End If
    
    FixedDays = GetFixedDaysFromParts(Year, Month, Day)
    
    ms = Statics.Calendar.GetTotalMilliseconds(Time)
    Set AddMonths = CorDateTime.FromMilliseconds(MilliSecondsPerDay * FixedDays + Modulus(ms, MilliSecondsPerDay))
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Seconds from the original Time.
'
' @param Time The time to add the specified number of Seconds to.
' @param Seconds The number of Seconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Seconds added.
'
Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As CorDateTime
    Set AddSeconds = Statics.Calendar.AddSeconds(Time, Seconds)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Weeks from the original Time.
'
' @param Time The time to add the specified number of Weeks to.
' @param Weeks The number of Weeks to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Weeks added.
'
Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As CorDateTime
    Set AddWeeks = Statics.Calendar.AddDays(Time, Weeks * DaysPerWeek)
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> this is the specified number of Years from the original Time.
'
' @param Time The time to add the specified number of Years to.
' @param Years The number of Years to be added to <i>Time</i>. This can be negative.
' @return Returns a <a href="CorDateTime.html">CorDateTime</a> with the specified number of Years added.
'
Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As CorDateTime
    Set AddYears = AddMonths(Time, Years * MonthsPerYear)
End Function

''
' Returns the day of the months for the specified date.
'
' @param Time The date inwhich to retrieve the day of the month.
' @return The day of the month for the specified date.
'
Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDateParts Time, DayPart, , , GetDayOfMonth
End Function

''
' Returns the day of the year for the specified date.
'
' @param Time The date to retrieve the day of the year from.
' @return The day of the year.
'
Public Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDateParts Time, DayOfTheYear, , , , GetDayOfYear
End Function

''
' Returns the <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> enum of the specified date.
'
' @param Time The date to retrieve the day of the week for.
' @return A <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> enum representing the day of the week.
' @remarks The <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> is Sunday and is 0-based. The Visual Basic
' enum vbDayOfWeek is 1-based.
'
Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = Statics.Calendar.GetDayOfWeek(Time)
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
    ValidateYearMonth Year, Month
    
    If Month = 12 Then
        If IsLeapYear(Year) Then
            GetDaysInMonth = 30
        Else
            GetDaysInMonth = 29
        End If
    ElseIf (Month And 1) = 1 Then
        GetDaysInMonth = 30
    Else
        GetDaysInMonth = 29
    End If
End Function

''
' Returns the number of days in the specified year.
'
' @param Year The year to retrieve the number of days of.
' @param Era The era the year exists in.
' @return The number of days in the year.
'
Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    If IsLeapYear(Year, Era) Then
        GetDaysInYear = 355
    Else
        GetDaysInYear = 354
    End If
End Function

''
' Returns the era that the specified time exists in.
'
' @param Time The time to find the containing era of.
' @return The era for the time.
'
Public Function GetEra(ByRef Time As Variant) As Long
    CorDateTime.GetDateTime Time
    GetEra = 1
End Function

''
' Returns the hour of the specified time.
'
' @param Time The time to retrieve the hour from.
' @return The hour portion of the time.
'
Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = Statics.Calendar.GetHour(Time)
End Function

''
' Returns the milliseconds of the specified time.
'
' @param Time The time to retrieve the milliseconds from.
' @return The millisecond portion of the time.
'
Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = Statics.Calendar.GetMilliseconds(Time)
End Function

''
' Returns the minute of the specified time.
'
' @param Time The time to retrieve the minute from.
' @return The minute portion of the time.
'
Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = Statics.Calendar.GetMinute(Time)
End Function

''
' Returns the month of the specified time.
'
' @param Time The time to retrieve the month from.
' @return The month portion of the <i>time</i>.
'
Public Function GetMonth(ByRef Time As Variant) As Long
    GetDateParts Time, MonthPart, , GetMonth
End Function

''
' Returns the number of months in the specified year.
'
' @param Year The year to get the number of months of.
' @param Era The era the year exists in.
' @return The number of months in the year.
'
Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    ValidateEra Era
    ValidateYear Year
    
    If Year = MaxYear Then
        GetMonthsInYear = MaxShortMonth
    Else
        GetMonthsInYear = MonthsPerYear
    End If
End Function

''
' Returns the second of the specified time.
'
' @param Time The time to retrieve the second from.
' @return The second portion of the time.
'
Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = Statics.Calendar.GetSecond(Time)
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
' Returns the year of the specified time.
'
' @param Time The time to retrieve the year from.
' @return The year portion of the time.
'
Public Function GetYear(ByRef Time As Variant) As Long
    GetDateParts Time, YearPart, GetYear
End Function

''
' Returns a boolean indicating if the specific day is a leap day.
'
' @param Year The year the day is in.
' @param Month The month the day is in.
' @param Day The day to check if is a leap day.
' @param Era The era the year is in.
' @return Boolean indication if the day is a leap day.
' @remarks <p>In every 30-year cycle that ends with a year that is evenly divisible by 30, the 2nd, 5th, 7th, 10th, 13th, 16th, 18th, 21st, 24th, 26th, and 29th years are leap years. A common year has 354 days and a leap year has 355 days.</p>
' <p>A leap day is a day that occurs only in a leap year. In the Hijri calendar, the 30th day of Zulhijjah is the only leap day.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'       <td>
'           <i>Year</i> is outside the range supported by the calendar.</br>
'           -or-</br>
'           <i>Month</i> is outside the range supported by the calendar.</br>
'           -or-</br>
'           <i>Day</i> is outside the range supported by the calendar.</br>
'           -or-</br>
'           <i>Era</i> is outside the range supported by the calendar.</br>
'       </td>
'   </tr>
' </table>
Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    Dim MaxDays As Long
    
    MaxDays = GetDaysInMonth(Year, Month)
    
    If Day < 1 Or Day > MaxDays Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, MaxDays))
        
    If Day = 30 And Month = 12 Then
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
'
Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    ValidateEra Era
    ValidateYearMonth Year, Month
    
    IsLeapMonth = False
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
    ValidateYear Year
    
    Select Case Year Mod 30
        Case 2, 5, 7, 10, 13, 16, 18, 21, 24, 26, 29
            IsLeapYear = True
    End Select
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
    Dim ms As Currency
    Dim DaysInMonth As Long
    
    DaysInMonth = GetDaysInMonth(Year, Month, Era)
    If Day < 1 Or Day > DaysInMonth Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Environment.GetResourceString(ArgumentOutOfRange_Day, DaysInMonth, Month))
    
    ms = GetFixedDaysFromParts(Year, Month, Day) * MilliSecondsPerDay
    
    If ms < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Message:="Invalid Year, Month, or Day for this calendar")
        
    Set ToDateTime = CorDateTime.FromMilliseconds(ms + Hour * MillisecondsPerHour + Minute * MillisecondsPerMinute + Second * MillisecondsPerSecond + Millisecond)
End Function

''
' Converts the specified year to a four-digit year by using the <a href="JulianCalendar.TwoDigitYearMax.get.html">TwoDigitYearMax</a> property to determine the appropriate century.
'
' @param A two-digit or four-digit integer that represents the year to convert.
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
    Dim y As Long
    
    If Year < 0 Or Year > MaxYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 9999))

    If Year < 100 Then
        y = Year Mod 100
        ToFourDigitYear = (mTwoDigitYearMax \ 100) * 100 + y
        
        If y > mTwoDigitYearMax Mod 100 Then
            ToFourDigitYear = ToFourDigitYear - 100
        End If
    Else
        ToFourDigitYear = Year
    End If
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
Private Sub VerifyWritable()
    If mIsReadOnly Then
        Error.InvalidOperation InvalidOperation_ReadOnly
    End If
End Sub

Private Sub ValidateEra(ByRef Era As Variant)
    Select Case VarType(Era)
        Case IsInteger(Err)
            If Era < Statics.Calendar.CurrentEra Or Era > Statics.HijriCalendar.HijriEra Then
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidEraValue), "Era")
            End If
        Case vbMissing
            ' do nothing
        Case Else
            Error.Argument Argument_IntegerRequired
    End Select
End Sub

Private Sub ValidateYear(ByVal Year As Long)
    If Year < 1 Or Year > MaxYear Then
        Throw Cor.NewArgumentOutOfRangeException("Year", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, MaxYear))
    End If
End Sub

Private Sub ValidateYearMonth(ByVal Year As Long, ByVal Month As Long)
    ValidateYear Year
    
    If (Month < 1) Or (Month > MonthsPerYear) Or (Year = MaxYear And Month > MaxShortMonth) Then
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_Range, 1, IIf(Year = MaxYear, 4, 12)))
    End If
End Sub

Private Function GetFixedDays(ByRef Time As Variant) As Long
    Dim TimeInMilliseconds As Currency
    
    TimeInMilliseconds = Statics.Calendar.GetTotalMilliseconds(Time)
    
    If TimeInMilliseconds < Statics.HijriCalendar.MinDateTime.TotalMilliseconds Or _
       TimeInMilliseconds > Statics.HijriCalendar.MaxDateTime.TotalMilliseconds Then
        Throw Cor.NewArgumentOutOfRangeException("Time", Environment.GetResourceString(ArgumentOutOfRange_CalendarRange, Statics.HijriCalendar.MinDateTime, Statics.HijriCalendar.MaxDateTime))
    End If
    
    GetFixedDays = Int(TimeInMilliseconds / MilliSecondsPerDay)
End Function

Private Function GetFixedDaysFromParts(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    Dim Result As Long
    Result = (Year - 1) * 354                   ' add days for each year, excluding first year.
    Result = Result + (Year * 11 + 3) \ 30      ' add additional days for leap years.
    Result = Result + (Month - 1) * 29          ' add the minimum of 29 days per month, excluding the first month.
    Result = Result + Month \ 2                 ' add the 1 day for each of the 30 day months.
    
    GetFixedDaysFromParts = Result + Day + EPOCH - 1 - mHijriAdjustment
End Function

Private Sub GetDateParts(ByRef Time As Variant, ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim FixedDays As Long
    FixedDays = GetFixedDays(Time)
    
    Year = (((FixedDays - EPOCH) * 30) + YearRoundingBias) \ DaysPer30Years
    If DatePart = YearPart Then
        Exit Sub
    End If
    
    DayOfYear = FixedDays - GetFixedDaysFromParts(Year, 1, 1) + 1
    If DatePart = DayOfTheYear Then
        Exit Sub
    End If
    
    If DayOfYear <> 355 Then
        Month = CorMath.Ceiling((DayOfYear - 30) / 29.5) + 1
    Else
        Month = 12
    End If
    
    If DatePart = MonthPart Then
        Exit Sub
    End If
    
    Day = FixedDays - GetFixedDaysFromParts(Year, Month, 1) + 1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mTwoDigitYearMax = DefaultTwoDigitYearMax
    mMonthDays = Cor.NewLongs(0, 30, 59, 89, 118, 148, 177, 207, 236, 266, 295, 325, 355)
    mIsReadOnly = True
End Sub

Friend Sub InitClone(ByVal Other As HijriCalendar)
    mHijriAdjustment = Other.HijriAdjustment
    mTwoDigitYearMax = Other.TwoDigitYearMax
    mIsReadOnly = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTwoDigitYearMax = .ReadProperty("TwoDigitYearMax")
        mHijriAdjustment = .ReadProperty("HijriAdjustment")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "TwoDigitYearMax", mTwoDigitYearMax
        .WriteProperty "HijriAdjustment", mHijriAdjustment
    End With
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

Private Property Get Calendar_DaysInYearBeforeMinSupportedYear() As Long
    Calendar_DaysInYearBeforeMinSupportedYear = DaysInYearBeforeMinSupportedYear
End Property

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



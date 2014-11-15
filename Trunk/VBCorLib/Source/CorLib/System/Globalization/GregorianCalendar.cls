VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GregorianCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: GregorianCalendar
'

''
' Provides methods for manipulating dates using the Gregorian calendar mathematics.
'
' @remarks The Gregorian calendar recognizes an era of BC (Before Christ) and
' AD (Anno Domino). The GregorianCalendar class only support the AD era (1).
' <p>The calendar supports the time from 01/01/0001 12:00:00AM to 12/31/9999 11:59:59PM.</p>
' <p>Only leap year is recognized and occurs every 4 years, except a year that is
' evenly divisible by 100. However, if it is divisible by 400 then it is a leap year.</p>
' <p>There are 12 months in the Gregorian calendar. The number of days range from 28 to
' 31, except on leap year when the shortest month is 29 days.</p>
'
' @see Calendar
' @see GregorianCalendarStatic
'
Option Explicit
Implements IObject
Implements Calendar


Public Enum GregorianCalendarTypes
    Localized = 1
    USEnglish = 2
    MiddleEastFrench = 9
    Arabic = 10
    TransliteratedEnglish = 11
    TransliteratedFrench = 12
End Enum


Private mCalendarType As GregorianCalendarTypes
Private mTwoDigitYearMax As Long



''
' Returns the calendar type that represents the language version.
'
' @return The calendar type that represents the language version.
' @remarks This value is only tracked with the calendar object. It
' does not apply languages to text values.
'
Public Property Get CalendarType() As GregorianCalendarTypes
    CalendarType = mCalendarType
End Property

''
' Sets the calendar type that represents the language version.
'
' @param RHS The calendar type to set this instance to.
' @remarks This value is only tracked with the calendar object. It
' does not apply languages to text values.
'
Public Property Let CalendarType(ByVal RHS As GregorianCalendarTypes)
    mCalendarType = RHS
End Property

''
' Returns an array of eras this calendar represents.
'
' @return An array containing the eras represented by this calendar.
' @remarks The Gregorian calendar recognizes only one era. The returned
' array will have one element containing the value 1.
'
Public Property Get Eras() As Long()
    Dim Ret(0) As Long
    Ret(0) = 1
    Eras = Ret
End Property

''
' Returns the maximum year to be created from a 2-digit year.
'
' @return The maximum year to be created.
' @remarks The property allows a 2-digit year (0 to 99) to be converted to
' a 4-digit year(0001 to 9999). The 2-digit year of 99 does not necessarily
' corrispond to the largest year to be created. For example, for a maximum
' year of 2029, the value 99 would convert to 1999 and 0 would convert to 2000.
' The years 29 would become 2029 and 30 would become 1930.
'
Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

''
' Sets the maximum year to be created from a 2-digit year.
'
' @param RHS The new 4-digit year maximum a 2-digit year is converted to.
' @remarks The property allows a 2-digit year (0 to 99) to be converted to
' a 4-digit year(0001 to 9999). The 2-digit year of 99 does not necessarily
' corrispond to the largest year to be created. For example, for a maximum
' year of 2029, the value 99 would convert to 1999 and 0 would convert to 2000.
' The years 29 would become 2029 and 30 would become 1930.
'
Public Property Let TwoDigitYearMax(ByVal RHS As Long)
    Const MinTwoDigitYear As Long = 100
    Const MaxTwoDigitYear As Long = 9999
    
    If RHS < MinTwoDigitYear Or RHS > MaxTwoDigitYear Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Message:=Resources.GetString(ArgumentOutOfRange_Range, MinTwoDigitYear, MaxTwoDigitYear))
    
    mTwoDigitYearMax = RHS
End Property

''
' Adds a number of days to the date provided using Gregorian math.
'
' @param Time This is a Date or cDateTime object to add a number of days to.
' @param Days The number of days to be added to <i>Time</i>.
' @return A <b>cDateTime</b> object with the new date.
' @remarks Using this method works exactly like using the <b>AddDays</b>
' method on the <b>cDateTime</b> objects.
' <p>A new cDateTime object is returned with the days added to the
' original <i>Time</i>. If a Date is passed in, a cDateTime object is
' still returned.</p>
' <p>A negative value can be supplied for <i>Days</i>.</p>
'
Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As cDateTime
    Set AddDays = cDateTime.GetcDateTime(Time).AddDays(Days)
End Function

''
' Adds a number of hours to the date provided.
'
' @param Time A Date or cDateTime object to add a number of hours to.
' @param Hours The number of hours to be added to the time portion of a date.
' @return Returns a <b>cDateTime</b> with the specified number of hours added.
'
Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As cDateTime
    Set AddHours = cDateTime.GetcDateTime(Time).AddHours(Hours)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Milliseconds from the original Time.
'
' @param Time The time to add the specified number of Milliseconds to.
' @param Milliseconds The number of Milliseconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Milliseconds added.
'
Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As cDateTime
    Set AddMilliseconds = cDateTime.GetcDateTime(Time).AddMilliseconds(Milliseconds)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Minutes from the original Time.
'
' @param Time The time to add the specified number of Minutes to.
' @param Minutes The number of Minutes to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Minutes added.
'
Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As cDateTime
    Set AddMinutes = cDateTime.GetcDateTime(Time).AddMinutes(Minutes)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Months from the original Time.
'
' @param Time The time to add the specified number of Months to.
' @param Months The number of Months to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Months added.
'
Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As cDateTime
    Set AddMonths = cDateTime.GetcDateTime(Time).AddMonths(Months)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Seconds from the original Time.
'
' @param Time The time to add the specified number of Seconds to.
' @param Seconds The number of Seconds to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Seconds added.
'
Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As cDateTime
    Set AddSeconds = cDateTime.GetcDateTime(Time).AddSeconds(Seconds)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Weeks from the original Time.
'
' @param Time The time to add the specified number of Weeks to.
' @param Weeks The number of Weeks to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Weeks added.
'
Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As cDateTime
    Set AddWeeks = AddDays(Time, Weeks * 7)
End Function

''
' Returns a <b>cDateTime</b> this is the specified number of Years from the original Time.
'
' @param Time The time to add the specified number of Years to.
' @param Years The number of Years to be added to <i>Time</i>. This can be negative.
' @return Returns a <b>cDateTime</b> with the specified number of Years added.
'
Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As cDateTime
    Set AddYears = cDateTime.GetcDateTime(Time).AddYears(Years)
End Function

''
' Returns the day of the months for the specified date.
'
' @param Time The date inwhich to retrieve the day of the month.
' @return The day of the month for the specified date.
'
Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    GetDayOfMonth = cDateTime.GetcDateTime(Time).Day
End Function

''
' Returns the <b>DayOfWeek</b> enum of the specified date.
'
' @param Time The date to retrieve the day of the week for.
' @return A <b>DayOfWeek</b> enum representing the day of the week.
' @remarks The <b>DayOfWeek</b> is Sunday and is 0-based. The Visual Basic
' enum vbDayOfWeek is 1-based.
'
Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = cDateTime.GetcDateTime(Time).DayOfWeek
End Function

''
' Returns the day of the year for the specified date.
'
' @param Time The date to retrieve the day of the year from.
' @return The day of the year.
'
Public Function GetDayOfYear(ByRef Time As Variant) As Long
    GetDayOfYear = cDateTime.GetcDateTime(Time).DayOfYear
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
    Call VerifyEra(Era)
    GetDaysInMonth = cDateTime.DaysInMonth(Year, Month)
End Function

''
' Returns the number of days in the specified year.
'
' @param Year The year to retrieve the number of days of.
' @param Era The era the year exists in.
' @return The number of days in the year.
' @remarks The gregorian calendar has 365 days in a normal year
' and 366 days in a leap year.
'
Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    If cDateTime.IsLeapYear(Year) Then
        GetDaysInYear = 366
    Else
        GetDaysInYear = 365
    End If
End Function

''
' Returns the era that the specified time exists in.
'
' @param Time The time to find the containing era of.
' @return The era for the time.
' @remarks The <b>GregorianCalendar</b> class only supports
' the current era of AD. This returns a constant of 1.
'
Public Function GetEra(ByRef Time As Variant) As Long
    Dim dt As cDateTime
    Set dt = cDateTime.GetcDateTime(Time)    ' verifies we have a date
    GetEra = 1
End Function

''
' Returns the hour of the specified time.
'
' @param Time The time to retrieve the hour from.
' @return The hour portion of the time.
'
Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = cDateTime.GetcDateTime(Time).Hour
End Function

''
' Returns the milliseconds of the specified time.
'
' @param Time The time to retrieve the milliseconds from.
' @return The millisecond portion of the time.
'
Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = cDateTime.GetcDateTime(Time).Millisecond
End Function

''
' Returns the minute of the specified time.
'
' @param Time The time to retrieve the minute from.
' @return The minute portion of the time.
'
Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = cDateTime.GetcDateTime(Time).Minute
End Function

''
' Returns the month of the specified time.
'
' @param Time The time to retrieve the month from.
' @return The month portion of the time.
'
Public Function GetMonth(ByRef Time As Variant) As Long
    GetMonth = cDateTime.GetcDateTime(Time).Month
End Function

''
' Returns the number of months in the specified year.
'
' @param Year The year to get the number of months of.
' @param Era The era the year exists in.
' @return The number of months in the year.
' @remarks The Gregorian calendar has 12 months in a year.
'
Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    GetMonthsInYear = 12
End Function

''
' Returns the second of the specified time.
'
' @param Time The time to retrieve the second from.
' @return The second portion of the time.
'
Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = cDateTime.GetcDateTime(Time).Second
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
    GetWeekOfYear = InternalGetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
End Function

''
' Returns the year of the specified time.
'
' @param Time The time to retrieve the year from.
' @return The year portion of the time.
'
Public Function GetYear(ByRef Time As Variant) As Long
    GetYear = cDateTime.GetcDateTime(Time).Year
End Function

''
' Returns a boolean indicating if the specific day is a leap day.
'
' @param Year The year the day is in.
' @param Month The month the day is in.
' @param Day The day to check if is a leap day.
' @param Era The era the year is in.
' @return Boolean indication if the day is a leap day.
' @remarks A leap day only exists in a leap year, such as Febuary 29th in the Gregorian calendar.
'
Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
    If Day < 1 Or Day > GetDaysInMonth(Year, Month) Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Day, Environment.GetResourceString(ArgumentOutOfRange_Range, 1, GetDaysInMonth(Year, Month)))
    
    If Month = 2 Then
        IsLeapDay = (Day = 29)
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
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
    
    IsLeapMonth = False
End Function

''
' Returns a boolean indicating if the specific year is a leap year.
'
' @param Year The year to check if is a leap year.
' @param Era The era the year is in.
' @return Boolean indication if the year is a leap year.
' @remarks A leap year in the Gregorian calendar occurs when the year is
' evenly divisible by 4, but not by 100, unless it is divisible by 400.
'
Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    
    IsLeapYear = cDateTime.IsLeapYear(Year)
End Function

''
' Returns a Gregorian cDateTime computed using the specific calendar rules.
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
Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByRef Era As Variant) As cDateTime
    Call VerifyEra(Era)
    Set ToDateTime = Cor.NewDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond)
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
    If Year < 0 Or Year > 9999 Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Year, Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 9999))

    If Year < 100 Then
        Dim y As Long
        y = Year Mod 100
        ToFourDigitYear = (mTwoDigitYearMax \ 100) * 100 + y
        If y > mTwoDigitYearMax Mod 100 Then ToFourDigitYear = ToFourDigitYear - 100
    Else
        ToFourDigitYear = Year
    End If
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal CalendarType As GregorianCalendarTypes)
    mCalendarType = CalendarType
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyEra(ByRef Era As Variant)
    If IsMissing(Era) Then Exit Sub
    Select Case VarType(Era)
        Case vbLong, vbInteger, vbByte
            If Era < 0 Or Era > 1 Then Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidEraValue), "Era")
        Case Else
            Throw Cor.NewInvalidCastException("An integer value is required.")
    End Select
End Sub

Private Sub VerifyYear(ByVal Year As Long)
    If Year < 1 Or Year > 9999 Then Throw Cor.NewArgumentOutOfRangeException("Year", Year, Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 9999))
End Sub

Private Sub VerifyMonth(ByVal Month As Long)
    If Month < 1 Or Month > 12 Then Throw Cor.NewArgumentOutOfRangeException("Month", Month, Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 12))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Calendar Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Calendar_AddDays(Time As Variant, ByVal Days As Long) As cDateTime
    Set Calendar_AddDays = AddDays(Time, Days)
End Function

Private Function Calendar_AddHours(Time As Variant, ByVal Hours As Long) As cDateTime
    Set Calendar_AddHours = AddHours(Time, Hours)
End Function

Private Function Calendar_AddMilliseconds(Time As Variant, ByVal Milliseconds As Double) As cDateTime
    Set Calendar_AddMilliseconds = AddMilliseconds(Time, Milliseconds)
End Function

Private Function Calendar_AddMinutes(Time As Variant, ByVal Minutes As Long) As cDateTime
    Set Calendar_AddMinutes = AddMinutes(Time, Minutes)
End Function

Private Function Calendar_AddMonths(Time As Variant, ByVal Months As Long) As cDateTime
    Set Calendar_AddMonths = AddMonths(Time, Months)
End Function

Private Function Calendar_AddSeconds(Time As Variant, ByVal Seconds As Long) As cDateTime
    Set Calendar_AddSeconds = AddSeconds(Time, Seconds)
End Function

Private Function Calendar_AddWeeks(Time As Variant, ByVal Weeks As Long) As cDateTime
    Set Calendar_AddWeeks = AddWeeks(Time, Weeks)
End Function

Private Function Calendar_AddYears(Time As Variant, ByVal Years As Long) As cDateTime
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

Private Function Calendar_ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional Era As Variant) As cDateTime
    Set Calendar_ToDateTime = ToDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond, Era)
End Function

Private Function Calendar_ToFourDigitYear(ByVal Year As Long) As Long
    Calendar_ToFourDigitYear = ToFourDigitYear(Year)
End Function

Private Function Calendar_ToString() As String
    Calendar_ToString = ToString
End Function

Private Property Let Calendar_TwoDigitYearMax(ByVal RHS As Long)
    TwoDigitYearMax = RHS
End Property

Private Property Get Calendar_TwoDigitYearMax() As Long
    Calendar_TwoDigitYearMax = TwoDigitYearMax
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mTwoDigitYearMax = GetCalendarLong(CAL_GREGORIAN, CAL_ITWODIGITYEARMAX)
    mCalendarType = GregorianCalendarTypes.Localized
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTwoDigitYearMax = .ReadProperty("TwoDigitYearMax")
        mCalendarType = .ReadProperty("CalendarType")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TwoDigitYearMax", mTwoDigitYearMax)
    Call PropBag.WriteProperty("CalendarType", mCalendarType)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: DateTimeFormatInfo
'

''
' Provides a repository of formatting information to be used in formatting
' cDateTime and Date values to match that of a specific locale.
'
' @remarks Formatting dates can vary greatly across multiple cultures. Each
' culture can provide formatting information that is specific to itself. A DateTimeFormatInfo
' object contains all the formatting information necessary to create a wide range of formatted
' dates and times. Also, this class provides a Format method to allow for directly
' formatting dates and times.<br>
'
' @see DateTimeFormatInfoStatic
' @see ICloneable
' @see IFormatProvider
' @include "..\..\Includes\DateTimeFormatInfo.txt"
Option Explicit
Implements IObject
Implements ICloneable
Implements IFormatProvider

Private Const PROP_ABBREVIATEDDAYNAMES      As String = "AbbreviatedDayNames"
Private Const PROP_ABBREVIATEDMONTHNAMES    As String = "AbbreviatedMonthNames"
Private Const PROP_AMDESIGNATOR             As String = "AMDesignator"
Private Const PROP_DATESEPARATOR            As String = "DateSeparator"
Private Const PROP_DAYNAMES                 As String = "DayNames"
Private Const PROP_FIRSTDAYOFWEEK           As String = "FirstDayOfWeek"
Private Const PROP_FULLDATETIMEPATTERN      As String = "FullDateTimePattern"
Private Const PROP_LONGDATEPATTERN          As String = "LongDatePattern"
Private Const PROP_LONGTIMEPATTERN          As String = "LongTimePattern"
Private Const PROP_MONTHDAYPATTERN          As String = "MonthDayPattern"
Private Const PROP_MONTHNAMES               As String = "MonthNames"
Private Const PROP_PMDESIGNATOR             As String = "PMDesignator"
Private Const PROP_SHORTDATEPATTERN         As String = "ShortDatePattern"
Private Const PROP_SHORTTIMEPATTERN         As String = "ShortTimePattern"
Private Const PROP_TIMESEPARATOR            As String = "TimeSeparator"
Private Const PROP_YEARMONTHPATTERN         As String = "YearMonthPattern"
Private Const PROP_CALENDARWEEKRULE         As String = "CalendarWeekRule"
Private Const PROP_CALENDAR                 As String = "Calendar"
Private Const PROP_ISREADONLY               As String = "IsReadOnly"

Private Const C_RFC1123Pattern                      As String = "ddd, dd MMM yyyy HH':'mm':'ss 'GMT'"
Private Const C_SORTABLEDATETIMEPATTERN             As String = "yyyy'-'MM'-'dd'T'HH':'mm':'ss"
Private Const C_UNIVERSALSORTABLEDATETIMEPATTERN    As String = "yyyy'-'MM'-'dd HH':'mm':'ss'Z'"

Private Const MinMonth As Long = 1
Private Const MaxMonth As Long = 13

Private Type PropsType
    AbbreviatedDayNames()   As String
    AbbreviatedMonthNames() As String
    AMDesignator            As String
    DateSeparator           As String
    DayNames()              As String
    FirstDayOfWeek          As DayOfWeek
    FullDateTimePattern     As Variant
    LongDatePattern         As String
    LongTimePattern         As String
    MonthDayPattern         As String
    MonthNames()            As String
    PMDesignator            As String
    ShortDatePattern        As String
    ShortTimePattern        As String
    TimeSeparator           As String
    YearMonthPattern        As String
    AllLongDatePatterns()   As String
    AllLongTimePatterns()   As String
    AllShortDatePatterns()  As String
    AllShortTimePatterns()  As String
    AllMonthDayPatterns()   As String
    Calendar                As Calendar
    CalendarWeekRule        As CalendarWeekRule
    
    IsReadOnly              As Boolean
End Type

Private mProps      As PropsType
Private mLoaded     As Boolean
Private mPattern    As WordBuffer


''
' Returns the definition of the first week of the calendar year.
'
' @return A <b>CalendarWeekRule</b> value.
'
Public Property Get CalendarWeekRule() As CalendarWeekRule
    VerifyLoaded
    CalendarWeekRule = mProps.CalendarWeekRule
End Property

''
' Sets the definition of the first week for the calendar year.
'
' @param RHS A <b>CalendarWeekRule</b> value.
'
Public Property Let CalendarWeekRule(ByVal RHS As CalendarWeekRule)
    VerifyWritable
    VerifyLoaded
    mProps.CalendarWeekRule = RHS
End Property

''
' Returns the <b>Calendar</b> associated with this <b>DateTimeFormatInfo</b>.
'
' @return A calendar object.
' @remarks If no calendar is associated then a <b>GregorianCalendar</b> is returned.
'
Public Property Get Calendar() As Calendar
    If mProps.Calendar Is Nothing Then Set mProps.Calendar = New GregorianCalendar
    Set Calendar = mProps.Calendar
End Property

''
' Sets the calendar associated with this <b>DateTimeFormatInfo</b>.
'
' @param RHS The calendar to associate with this object.
'
Public Property Set Calendar(ByVal Value As Calendar)
    VerifyWritable
    If Value Is Nothing Then _
        Throw Cor.NewArgumentNullException("Value")
    
    Set mProps.Calendar = Value
End Property

''
' Returns an array of abbreviated weekday names.
'
' @return A string array containing the abbreviated weekday names.
'
Public Property Get AbbreviatedDayNames() As String()
    VerifyLoaded
    AbbreviatedDayNames = mProps.AbbreviatedDayNames
End Property

''
' Sets the weekday names to be used by this FormatProvider.
'
' @param RHS An array of weekday names to be uses as abbreviated weekday names.
' @remarks The array must contain exactly 7 days. If the DateTimeFormatInfo
' is ReadOnly then an exception will be thrown.
'
Public Property Let AbbreviatedDayNames(ByRef RHS() As String)
    AssignDayNames RHS, mProps.AbbreviatedDayNames
End Property

''
' Returns an array of names for all the months.
'
' @return The array of names for all of the months.
'
Public Property Get AbbreviatedMonthNames() As String()
    VerifyLoaded
    AbbreviatedMonthNames = mProps.AbbreviatedMonthNames
End Property

''
' Sets the abbreviated month names to be used by the formatter.
'
' @param RHS The array of month names to be used.
' @remarks There must be exactly 13 months. Some cultures have 13 months in the
' year. For those cultures that only have 12, and empty string as the 13 is all that is needed.<br>
' If the formatter is ReadOnly, an exception will be thrown.
'
Public Property Let AbbreviatedMonthNames(ByRef RHS() As String)
    AssignMonthNames RHS, mProps.AbbreviatedMonthNames
End Property

''
' Returns the string that is displayed for times from 12:00:00 AM to 11:59:59 AM.
'
' @return The AM designator. Generally 'AM' is used.
'
Public Property Get AMDesignator() As String
    VerifyLoaded
    AMDesignator = mProps.AMDesignator
End Property

''
' Sets the string that is display for times from 12:00:00 AM to 11:59:59 AM.
'
' @param RHS The new AM designator.
' @remarks If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let AMDesignator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.AMDesignator = RHS
End Property

''
' Returns the string that is displayed when separating the components of a date,
' such as year, month, day.
'
' @return The date separator string. Generally this is '/'.
'
Public Property Get DateSeparator() As String
    VerifyLoaded
    DateSeparator = mProps.DateSeparator
End Property

''
' Sets the string that is display when separating the components of a date,
' such as year, month, day.
'
' @param RHS The new date separator. Generally this is '/'.
' @remarks If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let DateSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.DateSeparator = RHS
End Property

''
' Returns the names of the weekdays.
'
' @return An array containing the full names of the weekdays.
'
Public Property Get DayNames() As String()
    VerifyLoaded
    DayNames = mProps.DayNames
End Property

''
' Sets the names for the weekdays.
'
' @param RHS An array containing the weekday names to be used.
' @remarks There must be exactly 7 names. If this formatter is
' ReadOnly, an exception will be thrown.
'
Public Property Let DayNames(ByRef RHS() As String)
    VerifyWritable
    VerifyLoaded
    AssignDayNames RHS, mProps.DayNames
End Property

''
' Returns the first day of the week.
'
' @return The enum for the first day of the week.
' @remarks This DayOfWeek is 0 based, unlike VB's vbDayOfWeek which is 1 based.
'
Public Property Get FirstDayOfWeek() As DayOfWeek
    VerifyLoaded
    FirstDayOfWeek = mProps.FirstDayOfWeek
End Property

''
' Sets the first day of the week.
'
' @param RHS The enum for the first day of the week.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let FirstDayOfWeek(ByVal RHS As DayOfWeek)
    VerifyWritable
    VerifyLoaded
    mProps.FirstDayOfWeek = RHS
End Property

''
' Returns a formatting pattern that represents a full date and time associated
' with the 'F' formatting command.
'
' @return The full date and time pattern.
'
Public Property Get FullDateTimePattern() As String
    VerifyLoaded
    If IsEmpty(mProps.FullDateTimePattern) Then
        mProps.FullDateTimePattern = mProps.LongDatePattern & " " & mProps.LongTimePattern
    End If
    FullDateTimePattern = mProps.FullDateTimePattern
End Property

''
' Sets the formatting pattern used for a full date and time formatting,
' and is associated with the 'F' formatting command.
'
' @param RHS The new full date and time formatting pattern.
' @remark If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let FullDateTimePattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.FullDateTimePattern = RHS
End Property

''
' Returns if this formatter is ReadOnly.
'
' @return Indication if is ReadOnly or not.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

' for internal use only.
Friend Property Let IsReadOnly(ByVal RHS As Boolean)
    mProps.IsReadOnly = RHS
End Property

''
' Returns a pattern for formatting just the date portion and is
' associated with the 'D' formatting command.
'
' @return A date-only formatting.
'
Public Property Get LongDatePattern() As String
    VerifyLoaded
    LongDatePattern = mProps.LongDatePattern
End Property

''
' Sets a pattern used for formatting just the date portion and is
' associated with the 'D' formatting command.
'
' @param RHS The pattern used to format the date.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let LongDatePattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.LongDatePattern = RHS
    mProps.FullDateTimePattern = Empty
End Property

''
' Returns a pattern used in formatting the time portion and is
' associated with the 'T' formatting command.
'
' @return The time-only formatting pattern.
'
Public Property Get LongTimePattern() As String
    VerifyLoaded
    LongTimePattern = mProps.LongTimePattern
End Property

''
' Sets the pattern used to format the time part when using the
' 'T' formatting command.
'
' @param RHS The new formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let LongTimePattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.LongTimePattern = RHS
    mProps.FullDateTimePattern = Empty
End Property

''
' Returns the month/day pattern associated with either 'm' or 'M'
' formatting commands.
'
' @return The MonthDay formatting pattern.
'
Public Property Get MonthDayPattern() As String
    VerifyLoaded
    MonthDayPattern = mProps.MonthDayPattern
End Property

''
' Sets the month/day formatting pattern associated with either
' 'm' or 'M' formatting commands.
'
' @param RHS The new Month/Day pattern
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let MonthDayPattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.MonthDayPattern = RHS
End Property

''
' Returns the full names for the months of the year.
'
' @return An array containing the full month names.
'
Public Property Get MonthNames() As String()
    VerifyLoaded
    MonthNames = mProps.MonthNames
End Property

''
' Sets the full month names to be used by the formatter.
'
' @param RHS The array of month names.
' @remarks There must be exactly 13 months. Some cultures have 13 months
' in a calendar year. If a culture only has 12 months, then 13th month is empty.
' If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let MonthNames(ByRef RHS() As String)
    AssignMonthNames RHS, mProps.MonthNames
End Property

''
' Returns the string diplayed for times between 12:00:00 PM and 12:00:00 AM.
'
' @return The current PM designator. Generally this is 'PM'.
'
Public Property Get PMDesignator() As String
    VerifyLoaded
    PMDesignator = mProps.PMDesignator
End Property

''
' Sets the string display for times between 12:00:00 PM and 12:00:00 AM.
'
' @param RHS The new PM designator. Generally this is 'PM'.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let PMDesignator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.PMDesignator = RHS
End Property

''
' Returns the pattern as specified by RFC-1123 and associated with
' the 'r' or 'R' formatting commands.
'
' @return The RFC-1123 pattern <ddd, dd MMM yyyy HH':'mm':'ss 'GMT'>.
'
Public Property Get RFC1123Pattern() As String
    RFC1123Pattern = C_RFC1123Pattern
End Property

''
' Returns the pattern for short date formatting associated with the 'd' formatting command.
'
' @return The formatting pattern for a short date.
'
Public Property Get ShortDatePattern() As String
    VerifyLoaded
    ShortDatePattern = mProps.ShortDatePattern
End Property

''
' Sets the pattern used in formatting dates with the 'd' formatting command.
'
' @param RHS The new short date formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let ShortDatePattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.ShortDatePattern = RHS
End Property

''
' Returns the formatting pattern for a short time using the 't' formatting command.
'
' @return The short time formatting pattern.
'
Public Property Get ShortTimePattern() As String
    VerifyLoaded
    ShortTimePattern = mProps.ShortTimePattern
End Property

''
' Sets the formatting pattern for a short time using the 't' formatting command.
'
' @param RHS The new short time formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let ShortTimePattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.ShortDatePattern = RHS
End Property

''
' Returns a pattern that can be used in string sorting to produce an
' ascending sort with string based dates and times. This pattern is
' associated with the 's' formatting command.
'
' @return The string sortable pattern <yyyy'-'MM'-'dd'T'HH':'mm':'ss>.
' @remarks This pattern is the same across all cultures.
'
Public Property Get SortableDateTimePattern() As String
    SortableDateTimePattern = C_SORTABLEDATETIMEPATTERN
End Property

''
' Returns the separator of time components, such as hours, minutes, seconds.
'
' @return The time component separator.
'
Public Property Get TimeSeparator() As String
    VerifyLoaded
    TimeSeparator = mProps.TimeSeparator
End Property

''
' Set the separator for time components.
'
' @param RHS The new time component separator.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let TimeSeparator(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.TimeSeparator = RHS
End Property

''
' Returns a string sortable pattern that is associated with the 'u' or 'U'
' formatting commands.
'
' @return The UniversalSortableDateTimePattern <yyyy'-'MM'-'dd HH':'mm':'ss'Z'>.
' @remarks The pattern can be used to sort string representations of a date
' using the universal time designator "Z". This pattern is the same
' across all cultures.
'
Public Property Get UniversalSortableDateTimePattern() As String
    UniversalSortableDateTimePattern = C_UNIVERSALSORTABLEDATETIMEPATTERN
End Property

''
' Returns the pattern representing a YearMonth combination. This is
' associated with the 'y' or 'Y' formatting commands.
'
' @return The YearMonth formatting pattern.
'
Public Property Get YearMonthPattern() As String
    VerifyLoaded
    YearMonthPattern = mProps.YearMonthPattern
End Property

''
' Sets the pattern that represents a YearMonth combination. This is
' associated with the 'y' or 'Y' formatting commands.
'
' @param RHS The new YearMonth formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let YearMonthPattern(ByRef RHS As String)
    VerifyWritable
    VerifyLoaded
    mProps.YearMonthPattern = RHS
End Property

''
' Returns a clone of this instance. If this instance is ReadOnly,
' the clone will also be ReadOnly.
'
' @param The clones version of this instance.
'
Public Function Clone() As DateTimeFormatInfo
    Set Clone = New DateTimeFormatInfo
    Clone.CloneHelper mProps
End Function

''
' Returns the abbreviated name for the specified day-of-the-week.
'
' @param dow The Day-Of-Week to get the name of.
' @return The abbreviated version of the name for the day-of-week.
' @remarks DayOfWeek is 0 based. It is not identical to the vbDayOfweek,
' which is 1 based.
'
Public Function GetAbbreviatedDayName(ByVal dow As DayOfWeek) As String
    VerifyLoaded
    GetAbbreviatedDayName = mProps.AbbreviatedDayNames(dow)
End Function

''
' Returns the name of the Era specified.
'
' @param Era The era to retrieve the name for.
' @return The name of the era.
' @remarks Eras are identified by number. The gregorian calendar has two
' era, BC and AD. Only AD is recognized in this case by a number 1.
' <p>This feature is not fully implemented. It will only return 'AD'
' regardless of the culture.</p>
'
Public Function GetAbbreviatedEraName(ByVal Era As Long) As String
    VerifyLoaded
    Era = "AD"
End Function

''
' Retrieves the abbreviated name for a given month.
'
' @param Month The month number in the calendar starting at 1.
' @return The name of the specified month.
'
Public Function GetAbbreviatedMonthName(ByVal Month As Long) As String
    VerifyLoaded
    
    If Month < MinMonth Or Month > MaxMonth Then _
        Throw Cor.NewArgumentOutOfRangeException("Month", Month, Resources.GetString(ArgumentOutOfRange_Range, MinMonth, MaxMonth))
    
    GetAbbreviatedMonthName = mProps.AbbreviatedMonthNames(Month - 1)
End Function

''
' Returns all of the patterns for a specified formatting command. Or if no command
' is specified, then all of the patterns are returned.
'
' @param Format The formatting command to retrieve all the patterns for.
' @return An array of all the requested formatting patterns.
'
Public Function GetAllDateTimePatterns(Optional ByRef Format As Variant) As String()
    VerifyLoaded
    If IsMissing(Format) Then
        GetAllDateTimePatterns = CombineAllDateTimePatterns
    Else
        If VarType(Format) = vbString Then Format = Asc(Format)
        GetAllDateTimePatterns = GetSpecificDateTimePattern(Format)
    End If
End Function

''
' Returns the full name for the day of the week.
'
' @param dow The day of the week to get the name for.
' @return The name for the day of the week.
'
Public Function GetDayName(ByVal dow As DayOfWeek) As String
    VerifyLoaded
    GetDayName = mProps.DayNames(dow)
End Function

''
' Returns the numeric era value based on the name provided.
'
' @param eraName The name of the era to retrieve the numeric value for.
' @return The numeric value for the era.
' @remarks This method is not fully implemented. It will return 1
' for all era names and cultures.
'
Public Function GetEra(ByRef eraName As String) As Long
    VerifyLoaded
    ' need to update the culture table to handle multiple eras.
    GetEra = 1
End Function

''
' Returns the full name of the specified.
'
' @param Era The numerical era in which to retrieve the name for.
' @return The full era name.
' @remarks This method is not fully implements. It will return
' 'A.D.' for all numbers and cultures.
'
Public Function GetEraName(ByVal Era As Long) As String
    VerifyLoaded
    ' need to update the culture table to handle multiple eras.
    GetEraName = "A.D." ' mProps.Era
End Function

''
' Returns a format provider for the specified format type.
'
' @param FormatType The name of the formatting type requested.
' @return The formatting provider, or Nothing is none is available.
'
Public Function GetFormat(ByRef FormatType As String) As Object
    If StringUtils.EqualText(FormatType, "datetimeformatinfo") Then
        Set GetFormat = Me
    End If
End Function

''
' Returns the full name of the specified month.
'
' @param Month The month to retrieve the name for.
' @return The full name of the month.
' @remarks The valid range is 1 to 13.
'
Public Function GetMonthName(ByVal Month As Long) As String
    VerifyLoaded
    
    If Month < MinMonth Or Month > MaxMonth Then _
        Throw Cor.NewArgumentOutOfRangeException("Month", Month, Resources.GetString(ArgumentOutOfRange_Range, MinMonth, MaxMonth))
    
    GetMonthName = mProps.MonthNames(Month - 1)
End Function

''
' Formats a cDateTime object or Date value in the pattern specified by
' the formatting command. If no command is specified, then 'G' is assumed.
' If the command cannot be found, then a custom pattern is assumed.
'
' @param value The date to be formatted.
' @param fmt The formatting command or custom pattern to format the date to.
' @return The formatted date value.
' @remarks <p>Formatting can be accomplished by either specifying a formatting command,
' or entering a custom date format.<br><br>
' D - LongDatePattern<br>
' d - ShortDateTimePattern<br>
' F - FullDateTimePattern (Long Date and Long Time)<br>
' f - Full date and time (Long Date and Short Time)<br>
' G - General (Short Date and Long Time)<br>
' g - General (Short Date and Short Time)<br>
' m,M - MonthDayPattern<br>
' r,R - RFC1123Pattern<br>
' s - SortableDateTimePattern<br>
' T - LongTimePattern<br>
' t - ShortTimePattern<br>
' U - Full Date and Time (Long Date, Long Time) using universal time<br>
' u - UniversalSortableDateTimePattern<br>
' y,Y - YearMonthPattern<br>
'
Public Function Format(ByRef Value As Variant, Optional ByVal fmt As String) As String
    VerifyLoaded
    Format = CustomFormat(cDateTime.GetcDateTime(Value), GetPattern(fmt))
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
' @param value The value to compare equality to.
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
Friend Sub CloneHelper(ByRef props As PropsType)
    mProps = props
End Sub

Friend Sub Load(ByVal LCID As Long, ByVal UseUserOverride As Boolean, Optional ByVal Calendar As Calendar)
    If Not CultureTable.IsLoaded Then _
        Throw Cor.NewNotSupportedException("Cannot load DateTimeFormatInfo without culture table.")
    
    Set mProps.Calendar = Calendar
    LoadCommon LCID
    
    If UseUserOverride Then
        LoadUserOverride LCID
    Else
        LoadFromCultureTable LCID
    End If
    mLoaded = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUserOverride(ByVal LCID As Long)
    With mProps
        .CalendarWeekRule = GetLocaleLong(LCID, LOCALE_IFIRSTWEEKOFYEAR)
        .AMDesignator = GetLocaleString(LCID, LOCALE_S1159)
        .PMDesignator = GetLocaleString(LCID, LOCALE_S2359)
        .DateSeparator = GetLocaleString(LCID, LOCALE_SDATE)
        .TimeSeparator = GetLocaleString(LCID, LOCALE_STIME)
        .FirstDayOfWeek = GetLocaleLong(LCID, LOCALE_IFIRSTDAYOFWEEK)
        .LongDatePattern = GetLocaleString(LCID, LOCALE_SLONGDATE)
        .ShortDatePattern = GetLocaleString(LCID, LOCALE_SSHORTDATE)
        .LongTimePattern = GetLocaleString(LCID, LOCALE_STIMEFORMAT)
    End With
End Sub

Private Sub LoadFromCultureTable(ByVal LCID As Long)
    With mProps
        .CalendarWeekRule = CultureTable.GetNumber(LCID, IFIRSTWEEKOFYEAR)
        .AMDesignator = CultureTable.GetString(LCID, SAMDESIGNATOR)
        .PMDesignator = CultureTable.GetString(LCID, SPMDESIGNATOR)
        .DateSeparator = CultureTable.GetString(LCID, SDATESEPARATOR)
        .TimeSeparator = CultureTable.GetString(LCID, STIMESEPARATOR)
        .FirstDayOfWeek = CultureTable.GetNumber(LCID, IFIRSTDAYOFWEEK)
        .LongDatePattern = CultureTable.GetString(LCID, SLONGDATEPATTERN)
        .LongTimePattern = CultureTable.GetString(LCID, SLONGTIMEPATTERN)
        .ShortDatePattern = CultureTable.GetString(LCID, SSHORTDATEPATTERN)
    End With
End Sub

Private Sub LoadCommon(ByVal LCID As Long)
    With mProps
        .ShortTimePattern = CultureTable.GetString(LCID, SSHORTTIMEPATTERN)
        .AbbreviatedDayNames = CultureTable.GetStringArray(LCID, SABBREVIATEDDAYNAMES)
        .AbbreviatedMonthNames = CultureTable.GetStringArray(LCID, SABBREVIATEDMONTHNAMES)
        .DayNames = CultureTable.GetStringArray(LCID, SDAYNAMES)
        .MonthDayPattern = CultureTable.GetString(LCID, SMONTHDAYPATTERN)
        .MonthNames = CultureTable.GetStringArray(LCID, SMONTHNAMES)
        .YearMonthPattern = CultureTable.GetString(LCID, SYEARMONTHPATTERN)
        .AllLongDatePatterns = CultureTable.GetStringArray(LCID, SALLLONGDATEPATTERNS)
        .AllLongTimePatterns = CultureTable.GetStringArray(LCID, SALLLONGTIMEPATTERNS)
        .AllShortDatePatterns = CultureTable.GetStringArray(LCID, SALLSHORTDATEPATTERNS)
        .AllShortTimePatterns = CultureTable.GetStringArray(LCID, SALLSHORTTIMEPATTERNS)
        .AllMonthDayPatterns = CultureTable.GetStringArray(LCID, SALLMONTHDAYPATTERNS)
    End With
End Sub

Private Function CustomFormat(ByRef dt As cDateTime, ByRef Pattern As String) As String
    Dim cnt         As Long
    Dim Pos         As Long
    Dim sb          As StringBuilder
    Dim Day         As Long
    Dim Month       As Long
    Dim Year        As Long
    Dim YearMod100  As Long
    Dim Hour        As Long
    Dim HourMod12   As Long
    Dim Minute      As Long
    Dim Second      As Long
    Dim Designator  As String
    Dim ch          As Integer
    Dim Escaped     As Boolean
    Dim ZoneOffset  As TimeSpan
    Dim Percented   As Boolean
    Dim Quote       As Long
    
    mPattern.SA.pvData = StrPtr(Pattern)
    mPattern.SA.cElements = Len(Pattern)
    Set sb = DateTimeFormatInfo.StringBuilder   ' use a cached builder.
    sb.Length = 0
    
    Day = dt.Day
    Month = dt.Month
    Year = dt.Year
    YearMod100 = Year Mod 100
    Hour = dt.Hour
    HourMod12 = Hour Mod 12
    If HourMod12 = 0 Then HourMod12 = 12    ' we want 12:00:00 AM, not 00:00:00 AM
    Minute = dt.Minute
    Second = dt.Second
    Designator = IIf(Hour < 12, mProps.AMDesignator, mProps.PMDesignator)
    
    Do While Pos < mPattern.SA.cElements
        ch = mPattern.Data(Pos)
        cnt = GetRepeatCount(Pos)
        
        If Not Escaped Then
            If Quote <> 0 Then
                Select Case ch
                    Case vbBackslashChar:   Escaped = True
                    Case Quote:         Quote = 0
                    Case Else:          Call sb.AppendChar(ch)
                End Select
                cnt = 1
            
            Else
                Select Case ch
                    Case vbLowerDChar:          Call Append1to4CountValue(sb, cnt, Day, mProps.AbbreviatedDayNames(dt.DayOfWeek), mProps.DayNames(dt.DayOfWeek))
                    Case vbUpperMChar:          Call Append1to4CountValue(sb, cnt, Month, mProps.AbbreviatedMonthNames(Month - 1), mProps.MonthNames(Month - 1))
                    Case vbLowerHChar:          Call Append1or2DigitNumber(sb, cnt, HourMod12)
                    Case vbUpperHChar:          Call Append1or2DigitNumber(sb, cnt, Hour)
                    Case vbLowerMChar:          Call Append1or2DigitNumber(sb, cnt, Minute)
                    Case vbLowerSChar:          Call Append1or2DigitNumber(sb, cnt, Second)
                    Case vbColonChar:           Call sb.AppendString(mProps.TimeSeparator)
                    Case vbForwardSlashChar:    Call sb.AppendString(mProps.DateSeparator)
                    Case vbQuoteChar:     Quote = vbQuoteChar
                    Case vbSecondaryQuoteChar:     Quote = vbSecondaryQuoteChar
                    
                    Case vbLowerYChar
                        Select Case cnt
                            Case 1:         Call sb.Append(YearMod100)
                            Case 2:
                                If YearMod100 < 10 Then Call sb.AppendChar(vbZeroChar)
                                sb.Append YearMod100
                            Case Else
                                sb.AppendFormat "{0:d" & cnt & "}", Year
                        End Select
                        
                    Case vbLowerTChar
                        If cnt = 1 Then
                            sb.Append Designator, 0, 1
                        Else
                            sb.AppendString Designator
                        End If
                    
                    Case vbLowerZChar
                        If ZoneOffset Is Nothing Then Set ZoneOffset = TimeZone.CurrentTimeZone.GetUtcOffset(dt)
                        Select Case cnt
                            Case 1:         Call sb.AppendString(VBA.Format$(ZoneOffset.Hours, "0;-0"))
                            Case 2:         Call sb.AppendString(VBA.Format$(ZoneOffset.Hours, "00;-00"))
                            Case Else:      Call sb.AppendString(VBA.Format$(ZoneOffset.Hours, "00;-00") & ":" & VBA.Format$(ZoneOffset.Minutes, "00"))
                        End Select
                    
                    Case vbPercentChar
                        If Percented Then Call FormatError
                        cnt = 1
                    
                    Case vbBackslashChar
                        Escaped = True
                        cnt = 1
                    
                    Case vbLowerFChar
                        Dim secfrac As Long
                        If cnt > 7 Then Call FormatError
                        secfrac = GetSecondsFraction(dt.Ticks)
                        secfrac = (secfrac \ CLng(10 ^ (7 - cnt)))
                        Call sb.AppendFormat("{0:g" & cnt & "}", secfrac)
                        
                    Case Else:              Call sb.AppendChar(ch)
                End Select
            End If
        Else
            Call sb.AppendChar(ch)
            Escaped = False
            cnt = 1
        End If
        Percented = (ch = vbPercentChar)
        Pos = Pos + cnt
    Loop
    
    If Quote <> 0 Then _
        Throw Cor.NewFormatException("A matching quote was not found in the format string.")
    
    CustomFormat = sb.ToString
End Function

Private Function GetSecondsFraction(ByRef Ticks As Variant) As Long
    Const TICKS_PER_SECOND As Long = 10000000
    GetSecondsFraction = Ticks - (TICKS_PER_SECOND * Fix(Ticks / TICKS_PER_SECOND))
End Function

Private Sub FormatError()
    Throw Cor.NewFormatException("The string was in the incorrect format.")
End Sub

Private Sub Append1to4CountValue(ByRef sb As StringBuilder, ByVal cnt As Long, ByVal Value As Long, ByRef ThreeCountValue As String, ByRef FourCountValue As String)
    Select Case cnt
        Case 1
            Call sb.Append(Value)
        Case 2
            If Value < 10 Then Call sb.AppendChar(vbZeroChar)
            Call sb.Append(Value)
        Case 3
            Call sb.AppendString(ThreeCountValue)
        Case Else
            Call sb.AppendString(FourCountValue)
    End Select
End Sub

Private Sub Append1or2DigitNumber(ByRef sb As StringBuilder, ByVal cnt As Long, ByVal Value As Long)
    If cnt = 1 Then
        Call sb.Append(Value)
    Else
        If Value < 10 Then Call sb.AppendChar(vbZeroChar)
        Call sb.Append(Value)
    End If
End Sub

''
' Given an index in the mPattern, the counts the number of times
' the character at the specific index repeats starting at that index.
'
' @param Index The starting location in mPatterns to be counting repeate characters.
' @return The number of time the character repeated starting at the index.
'
Private Function GetRepeatCount(ByVal Index As Long) As Long
    Dim i As Long
    Dim ch As Integer
    
    ch = mPattern.Data(Index)
    i = Index + 1
    Do While i < mPattern.SA.cElements
        If mPattern.Data(i) <> ch Then Exit Do
        i = i + 1
    Loop
    GetRepeatCount = i - Index
End Function

Private Function GetPattern(ByVal fmt As String) As String
    If Len(fmt) = 0 Then fmt = "G"
    If Len(fmt) = 1 Then
        Select Case Asc(fmt)
            Case vbLowerDChar:               GetPattern = mProps.ShortDatePattern
            Case vbUpperDChar:               GetPattern = mProps.LongDatePattern
            Case vbLowerTChar:               GetPattern = mProps.ShortTimePattern
            Case vbUpperTChar:               GetPattern = mProps.LongTimePattern
            Case vbLowerFChar:               GetPattern = mProps.LongDatePattern & " " & mProps.ShortTimePattern
            Case vbUpperFChar:               GetPattern = Me.FullDateTimePattern
            Case vbLowerGChar:               GetPattern = mProps.ShortDatePattern & " " & mProps.ShortTimePattern
            Case vbUpperGChar:               GetPattern = mProps.ShortDatePattern & " " & mProps.LongTimePattern
            Case vbLowerMChar, vbUpperMChar:     GetPattern = mProps.MonthDayPattern
            Case vbLowerRChar, vbUpperRChar:     GetPattern = C_RFC1123Pattern
            Case vbLowerSChar:               GetPattern = C_SORTABLEDATETIMEPATTERN
            Case vbLowerUChar:               GetPattern = C_UNIVERSALSORTABLEDATETIMEPATTERN
            Case vbUpperUChar:               GetPattern = Me.FullDateTimePattern
            Case vbLowerYChar, vbUpperYChar:     GetPattern = mProps.YearMonthPattern
            Case Else
                Throw Cor.NewFormatException("Invalid format specifier.")
        End Select
    Else
        GetPattern = fmt
    End If
End Function

''
' Returns a list of date and time patterns avaiable.
Private Function CombineAllDateTimePatterns() As String()
    Const FORMATS As String = "dDfFgGmMrRstTUuYy"
    mPattern.SA.pvData = StrPtr(FORMATS)
    
    Dim List As ArrayList
    Set List = New ArrayList
    
    Dim i As Long
    For i = 0 To Len(FORMATS) - 1
        Call List.AddRange(GetSpecificDateTimePattern(mPattern.Data(i)))
    Next i
    
    Dim Ret() As String
    ReDim Ret(List.Count - 1)
    Call List.CopyTo(Ret)
    CombineAllDateTimePatterns = Ret
End Function

''
' Get the Date&Time pattern(s) based on the requested type.
'
' @param Pattern The requested pattern.
' @return An array of all Date&Time patterns of the requested type.
'
Private Function GetSpecificDateTimePattern(ByVal Pattern As Long) As String()
    Dim Ret() As String
    
    ' We'll redim here and take the hit if the client
    ' requests one of the multi-pattern types.
    ' It will be rare that a multi-pattern is requested.
    ReDim Ret(0)
    Select Case Pattern
        Case vbLowerDChar
            Ret(0) = mProps.ShortDatePattern
        Case vbUpperDChar
            Ret(0) = mProps.LongDatePattern
        Case vbLowerTChar
            Ret = mProps.AllShortTimePatterns
        Case vbUpperTChar
            Ret(0) = mProps.LongTimePattern
        Case vbUpperFChar, vbUpperUChar
            Ret(0) = Me.FullDateTimePattern
        Case vbLowerFChar
            Ret = CreateDateTimePatterns(mProps.AllLongDatePatterns, mProps.AllShortTimePatterns)
        Case vbLowerGChar
            Ret = CreateDateTimePatterns(mProps.AllShortDatePatterns, mProps.AllShortTimePatterns)
        Case vbUpperGChar
            Ret = CreateDateTimePatterns(mProps.AllShortDatePatterns, mProps.AllLongTimePatterns)
        Case vbUpperMChar, vbLowerMChar
            Ret(0) = mProps.MonthDayPattern
        Case vbUpperRChar, vbLowerRChar
            Ret(0) = C_RFC1123Pattern
        Case vbLowerSChar
            Ret(0) = C_SORTABLEDATETIMEPATTERN
        Case vbLowerUChar
            Ret(0) = C_UNIVERSALSORTABLEDATETIMEPATTERN
        Case vbLowerYChar, vbUpperYChar
            Ret(0) = mProps.YearMonthPattern
        Case Else
            Throw Cor.NewArgumentException("Invalid format specifier.")
    End Select
    GetSpecificDateTimePattern = Ret
End Function

''
' This creates Date&Time patterns by combining all the date patterns
' with all time patterns in all possible combinations.
'
' @param DatePatterns The date patterns to combine with time patterns.
' @param TimePatterns The time patterns to combine with date patterns.
' @return An array of Date&Time pattern combinations.
'
Private Function CreateDateTimePatterns(ByRef DatePatterns() As String, ByRef TimePatterns() As String) As String()
    Dim TimePatternsUpperBound As Long
    TimePatternsUpperBound = UBound(TimePatterns)
    
    Dim DatePatternsUpperBound As Long
    DatePatternsUpperBound = UBound(DatePatterns)
    
    Dim Ret() As String
    ReDim Ret(0 To (DatePatternsUpperBound + 1) * (TimePatternsUpperBound + 1) - 1)
    
    Dim DatePatternIndex As Long
    For DatePatternIndex = 0 To DatePatternsUpperBound
        Dim TimePatternIndex As Long
        For TimePatternIndex = 0 To TimePatternsUpperBound
            Dim i As Long
            Ret(i) = DatePatterns(DatePatternIndex) & " " & TimePatterns(TimePatternIndex)
            i = i + 1
        Next TimePatternIndex
    Next DatePatternIndex
    CreateDateTimePatterns = Ret
End Function

Private Sub VerifyWritable()
    If mProps.IsReadOnly Then Throw Cor.NewInvalidOperationException(Environment.GetResourceString(InvalidOperation_ReadOnly))
End Sub

Private Sub ReadStringArray(ByRef a() As String, ByVal Name As String, ByVal bag As PropertyBag)
    With bag
        Dim Count As Long
        Count = .ReadProperty(Name & "_Count", 0)
        a = cArray.CreateInstance(vbString, Count)
        
        Dim i As Long
        For i = 0 To Count - 1
            a(i) = .ReadProperty(Name & "_" & i)
        Next i
    End With
End Sub

Private Sub WriteStringArray(ByRef a() As String, ByVal Name As String, ByVal bag As PropertyBag)
    With bag
        Call .WriteProperty(Name & "_Count", UBound(a) + 1)
        Dim i As Long
        For i = 0 To UBound(a)
            Call .WriteProperty(Name & "_" & i, a(i))
        Next i
    End With
End Sub

Private Sub VerifyLoaded()
    If mLoaded Then Exit Sub
    Call Load(INVARIANT_LCID, False)
End Sub

Private Sub AssignDayNames(ByRef Source() As String, ByRef DayNames() As String)
    Call VerifyWritable
    Call VerifyLoaded

    If cArray.GetRank(Source) <> 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimNotSupported))
    If cArray.GetLength(Source) <> 7 Then _
        Throw Cor.NewArgumentException("Array must contain exactly 7 elements.")
    
    ReDim DayNames(0 To 6)
    Call cArray.CopyEx(Source, LBound(Source), DayNames, 0, 7)
End Sub

Private Sub AssignMonthNames(ByRef Source() As String, ByRef MonthNames() As String)
    Call VerifyWritable
    Call VerifyLoaded
    If cArray.GetRank(Source) <> 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimNotSupported))
    If cArray.GetLength(Source) <> 13 Then _
        Throw Cor.NewArgumentException("Array must contain exactly 13 elements.")
        
    ReDim MonthNames(0 To 12)
    Call cArray.CopyEx(Source, LBound(Source), MonthNames, 0, 13)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWordBuffer(mPattern, 0, &H7FFFFFFF)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With mProps
        Set .Calendar = PropBag.ReadProperty(PROP_CALENDAR, Nothing)
        .CalendarWeekRule = PropBag.ReadProperty(PROP_CALENDARWEEKRULE)
        .AMDesignator = PropBag.ReadProperty(PROP_AMDESIGNATOR)
        .DateSeparator = PropBag.ReadProperty(PROP_DATESEPARATOR)
        .FirstDayOfWeek = PropBag.ReadProperty(PROP_FIRSTDAYOFWEEK)
        .FullDateTimePattern = PropBag.ReadProperty(PROP_FULLDATETIMEPATTERN)
        .LongDatePattern = PropBag.ReadProperty(PROP_LONGDATEPATTERN)
        .LongTimePattern = PropBag.ReadProperty(PROP_LONGTIMEPATTERN)
        .MonthDayPattern = PropBag.ReadProperty(PROP_MONTHDAYPATTERN)
        .PMDesignator = PropBag.ReadProperty(PROP_PMDESIGNATOR)
        .ShortDatePattern = PropBag.ReadProperty(PROP_SHORTDATEPATTERN)
        .ShortTimePattern = PropBag.ReadProperty(PROP_SHORTTIMEPATTERN)
        .TimeSeparator = PropBag.ReadProperty(PROP_TIMESEPARATOR)
        .YearMonthPattern = PropBag.ReadProperty(PROP_YEARMONTHPATTERN)
        .IsReadOnly = PropBag.ReadProperty(PROP_ISREADONLY)
        Call ReadStringArray(.AbbreviatedDayNames, PROP_ABBREVIATEDDAYNAMES, PropBag)
        Call ReadStringArray(.AbbreviatedMonthNames, PROP_ABBREVIATEDMONTHNAMES, PropBag)
        Call ReadStringArray(.MonthNames, PROP_MONTHNAMES, PropBag)
        Call ReadStringArray(.DayNames, PROP_DAYNAMES, PropBag)
    End With
    mLoaded = True
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With mProps
        Call PropBag.WriteProperty(PROP_CALENDAR, .Calendar)
        Call PropBag.WriteProperty(PROP_CALENDARWEEKRULE, .CalendarWeekRule)
        Call WriteStringArray(.AbbreviatedDayNames, PROP_ABBREVIATEDDAYNAMES, PropBag)
        Call WriteStringArray(.AbbreviatedMonthNames, PROP_ABBREVIATEDMONTHNAMES, PropBag)
        Call PropBag.WriteProperty(PROP_AMDESIGNATOR, .AMDesignator)
        Call PropBag.WriteProperty(PROP_DATESEPARATOR, .DateSeparator)
        Call WriteStringArray(.DayNames, PROP_DAYNAMES, PropBag)
        Call PropBag.WriteProperty(PROP_FIRSTDAYOFWEEK, .FirstDayOfWeek)
        Call PropBag.WriteProperty(PROP_FULLDATETIMEPATTERN, .FullDateTimePattern)
        Call PropBag.WriteProperty(PROP_LONGDATEPATTERN, .LongDatePattern)
        Call PropBag.WriteProperty(PROP_LONGTIMEPATTERN, .LongTimePattern)
        Call PropBag.WriteProperty(PROP_MONTHDAYPATTERN, .MonthDayPattern)
        Call WriteStringArray(.MonthNames, PROP_MONTHNAMES, PropBag)
        Call PropBag.WriteProperty(PROP_PMDESIGNATOR, .PMDesignator)
        Call PropBag.WriteProperty(PROP_SHORTDATEPATTERN, .ShortDatePattern)
        Call PropBag.WriteProperty(PROP_SHORTTIMEPATTERN, .ShortTimePattern)
        Call PropBag.WriteProperty(PROP_TIMESEPARATOR, .TimeSeparator)
        Call PropBag.WriteProperty(PROP_YEARMONTHPATTERN, .YearMonthPattern)
        Call PropBag.WriteProperty(PROP_ISREADONLY, .IsReadOnly)
    End With
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function

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
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: DateTimeFormatInfo
'

''
' Provides a repository of formatting information to be used in formatting
' CorDateTime and Date values to match that of a specific locale.
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

Private Const PropAbbreviatedDayNames   As String = "AbbreviatedDayNames"
Private Const PropAbbreviatedMonthNames As String = "AbbreviatedMonthNames"
Private Const PropAMDesignator          As String = "AMDesignator"
Private Const PropDateSeparator         As String = "DateSeparator"
Private Const PropDayNames              As String = "DayNames"
Private Const PropFirstDayOfWeek        As String = "FirstDayOfWeek"
Private Const PropFullDateTimePattern   As String = "FullDateTimePattern"
Private Const PropLongDatePattern       As String = "LongDatePattern"
Private Const PropLongTimePattern       As String = "LongTimePattern"
Private Const PropMonthDayPattern       As String = "MonthDayPattern"
Private Const PropMonthNames            As String = "MonthNames"
Private Const PropPMDesignator          As String = "PMDesignator"
Private Const PropShortDatePattern      As String = "ShortDatePattern"
Private Const PropShortTimePattern      As String = "ShortTimePattern"
Private Const PropTimeSeparator         As String = "TimeSeparator"
Private Const PropYearMonthPattern      As String = "YearMonthPattern"
Private Const PropCalendarWeekRule      As String = "CalendarWeekRule"
Private Const PropCalendar              As String = "Calendar"
Private Const PropIsReadOnly            As String = "IsReadOnly"

Private Const PatternRFC1123                    As String = "ddd, dd MMM yyyy HH':'mm':'ss 'GMT'"
Private Const PatternSortableDateTime           As String = "yyyy'-'MM'-'dd'T'HH':'mm':'ss"
Private Const PatternUniversalSortableDateTime  As String = "yyyy'-'MM'-'dd HH':'mm':'ss'Z'"

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
Private mPattern    As CharBuffer


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
' @param Value A <b>CalendarWeekRule</b> value.
'
Public Property Let CalendarWeekRule(ByVal Value As CalendarWeekRule)
    VerifyWritable
    VerifyLoaded
    mProps.CalendarWeekRule = Value
End Property

''
' Returns the <b>Calendar</b> associated with this <b>DateTimeFormatInfo</b>.
'
' @return A calendar object.
' @remarks If no calendar is associated then a <b>GregorianCalendar</b> is returned.
'
Public Property Get Calendar() As Calendar
    If mProps.Calendar Is Nothing Then
        Set mProps.Calendar = New GregorianCalendar
    End If
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
        Error.ArgumentNull "Value"
    
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
' @param Value An array of weekday names to be uses as abbreviated weekday names.
' @remarks The array must contain exactly 7 days. If the DateTimeFormatInfo
' is ReadOnly then an exception will be thrown.
'
Public Property Let AbbreviatedDayNames(ByRef Value() As String)
    mProps.AbbreviatedDayNames = CopyArray(Value, 7)
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
' @param Value The array of month names to be used.
' @remarks There must be exactly 13 months. Some cultures have 13 months in the
' year. For those cultures that only have 12, and empty string as the 13 is all that is needed.<br>
' If the formatter is ReadOnly, an exception will be thrown.
'
Public Property Let AbbreviatedMonthNames(ByRef Value() As String)
    mProps.AbbreviatedMonthNames = CopyArray(Value, 13)
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
' @param Value The new AM designator.
' @remarks If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let AMDesignator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.AMDesignator = Value
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
' @param Value The new date separator. Generally this is '/'.
' @remarks If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let DateSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.DateSeparator = Value
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
' @param Value An array containing the weekday names to be used.
' @remarks There must be exactly 7 names. If this formatter is
' ReadOnly, an exception will be thrown.
'
Public Property Let DayNames(ByRef Value() As String)
    mProps.DayNames = CopyArray(Value, 7)
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
' @param Value The enum for the first day of the week.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let FirstDayOfWeek(ByVal Value As DayOfWeek)
    VerifyWritable
    VerifyLoaded
    mProps.FirstDayOfWeek = Value
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
' @param Value The new full date and time formatting pattern.
' @remark If this formatter is ReadOnly, an exception will be thrown.
'
Public Property Let FullDateTimePattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.FullDateTimePattern = Value
End Property

''
' Returns if this formatter is ReadOnly.
'
' @return Indication if is ReadOnly or not.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

Friend Sub MarkReadOnly()
    mProps.IsReadOnly = True
End Sub

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
' @param Value The pattern used to format the date.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let LongDatePattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.LongDatePattern = Value
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
' @param Value The new formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let LongTimePattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.LongTimePattern = Value
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
' @param Value The new Month/Day pattern
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let MonthDayPattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.MonthDayPattern = Value
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
' @param Value The array of month names.
' @remarks There must be exactly 13 months. Some cultures have 13 months
' in a calendar year. If a culture only has 12 months, then 13th month is empty.
' If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let MonthNames(ByRef Value() As String)
    mProps.MonthNames = CopyArray(Value, 13)
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
' @param Value The new PM designator. Generally this is 'PM'.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let PMDesignator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.PMDesignator = Value
End Property

''
' Returns the pattern as specified by RFC-1123 and associated with
' the 'r' or 'R' formatting commands.
'
' @return The RFC-1123 pattern <ddd, dd MMM yyyy HH':'mm':'ss 'GMT'>.
'
Public Property Get RFC1123Pattern() As String
    RFC1123Pattern = PatternRFC1123
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
' @param Value The new short date formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let ShortDatePattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.ShortDatePattern = Value
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
' @param Value The new short time formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let ShortTimePattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.ShortDatePattern = Value
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
    SortableDateTimePattern = PatternSortableDateTime
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
' @param Value The new time component separator.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let TimeSeparator(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.TimeSeparator = Value
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
    UniversalSortableDateTimePattern = PatternUniversalSortableDateTime
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
' @param Value The new YearMonth formatting pattern.
' @remarks If this formatter is ReadOnly, an exception is thrown.
'
Public Property Let YearMonthPattern(ByRef Value As String)
    VerifyWritable
    VerifyLoaded
    mProps.YearMonthPattern = Value
End Property

''
' Returns a clone of this instance. If this instance is ReadOnly,
' the clone will also be ReadOnly.
'
' @param The clones version of this instance.
'
Public Function Clone() As DateTimeFormatInfo
    Set Clone = New DateTimeFormatInfo
    Clone.InitClone mProps
End Function

''
' Returns the abbreviated name for the specified day-of-the-week.
'
' @param DayOfWeek The Day-Of-Week to get the name of.
' @return The abbreviated version of the name for the day-of-week.
' @remarks DayOfWeek is 0 based. It is not identical to the vbDayOfweek,
' which is 1 based.
'
Public Function GetAbbreviatedDayName(ByVal DayOfWeek As DayOfWeek) As String
    If DayOfWeek < 0 Or DayOfWeek > 6 Then _
        Throw Cor.NewArgumentOutOfRangeException("DayOfWeek", Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 6))

    VerifyLoaded
    GetAbbreviatedDayName = mProps.AbbreviatedDayNames(DayOfWeek)
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
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_Range, MinMonth, MaxMonth))
    
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
    ElseIf VarType(Format) = vbString Then
        GetAllDateTimePatterns = GetSpecificDateTimePattern(AscW(Format))
    Else
        GetAllDateTimePatterns = GetSpecificDateTimePattern(CLng(Format))
    End If
End Function

''
' Returns the full name for the day of the week.
'
' @param DayOfWeek The day of the week to get the name for.
' @return The name for the day of the week.
'
Public Function GetDayName(ByVal DayOfWeek As DayOfWeek) As String
    VerifyLoaded
    GetDayName = mProps.DayNames(DayOfWeek)
End Function

''
' Returns the numeric era value based on the name provided.
'
' @param eraName The name of the era to retrieve the numeric value for.
' @return The numeric value for the era.
' @remarks This method is not fully implemented. It will return 1
' for all era names and cultures.
'
Public Function GetEra(ByRef EraName As String) As Long
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
    If CorString.Equals(FormatType, corDateTimeFormatInfo, OrdinalIgnoreCase) Then
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
        Throw Cor.NewArgumentOutOfRangeException("Month", Environment.GetResourceString(ArgumentOutOfRange_Range, MinMonth, MaxMonth))
    
    GetMonthName = mProps.MonthNames(Month - 1)
End Function

''
' Formats a CorDateTime object or Date value in the pattern specified by
' the formatting command. If no command is specified, then 'G' is assumed.
' If the command cannot be found, then a custom pattern is assumed.
'
' @param Value The date to be formatted.
' @param FormatString The formatting command or custom pattern to format the date to.
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
Friend Function Format(ByVal Value As CorDateTime, Optional ByRef FormatString As String) As String
    VerifyLoaded
    Format = CustomFormat(Value, GetPattern(FormatString))
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
' @param value The value to compare equality to.
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
Friend Sub InitClone(ByRef Props As PropsType)
    mProps = Props
End Sub

Friend Sub Init(ByVal LCID As Long, ByVal UseUserOverride As Boolean, Optional ByVal Calendar As Calendar)
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

Friend Sub InitCultureInfo(ByVal Info As CultureInfo)
    Init Info.LCID, Info.UseUserOverride, Info.Calendar
    If Info.IsReadOnly Then
        MarkReadOnly
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
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

Private Function CustomFormat(ByRef dt As CorDateTime, ByRef Pattern As String) As String
    Dim Count       As Long
    Dim pos         As Long
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
    Dim Ch          As Integer
    Dim Escaped     As Boolean
    Dim ZoneOffset  As TimeSpan
    Dim Percented   As Boolean
    Dim Quote       As Long
    
    InitChars mPattern, Pattern
    Set sb = StringBuilderCache.Acquire
    
    On Error GoTo Catch
    
    dt.GetDateParts DatePartPrecision.Complete, Year, Month, Day
    
    YearMod100 = Year Mod 100
    Hour = dt.Hour
    HourMod12 = Hour Mod 12
    If HourMod12 = 0 Then
        HourMod12 = 12    ' we want 12:00:00 AM, not 00:00:00 AM
    End If
    
    Minute = dt.Minute
    Second = dt.Second
    
    If Hour < 12 Then
        Designator = mProps.AMDesignator
    Else
        Designator = mProps.PMDesignator
    End If
    
    Do While pos < mPattern.Buffer.cElements
        Ch = mPattern.Chars(pos)
        Count = GetRepeatCount(pos)
        
        If Not Escaped Then
            If Quote <> 0 Then
                Select Case Ch
                    Case corBackslashChar
                        Escaped = True
                    Case Quote
                        Quote = 0
                    Case Else
                        sb.AppendChar Ch
                End Select
                Count = 1
            
            Else
                Select Case Ch
                    Case corLowerDChar
                        Append1to4CountValue sb, Count, Day, mProps.AbbreviatedDayNames(dt.DayOfWeek), mProps.DayNames(dt.DayOfWeek)
                    Case corUpperMChar
                        Append1to4CountValue sb, Count, Month, mProps.AbbreviatedMonthNames(Month - 1), mProps.MonthNames(Month - 1)
                    Case corLowerHChar
                        Append1or2DigitNumber sb, Count, HourMod12
                    Case corUpperHChar
                        Append1or2DigitNumber sb, Count, Hour
                    Case corLowerMChar
                        Append1or2DigitNumber sb, Count, Minute
                    Case corLowerSChar
                        Append1or2DigitNumber sb, Count, Second
                    Case corColonChar
                        sb.AppendString mProps.TimeSeparator
                    Case corForwardSlashChar
                        sb.AppendString mProps.DateSeparator
                    Case corQuoteChar
                        Quote = corQuoteChar
                    Case corSecondaryQuoteChar
                        Quote = corSecondaryQuoteChar
                    Case corLowerYChar
                        Select Case Count
                            Case 1
                                sb.Append YearMod100
                            Case 2
                                If YearMod100 < 10 Then
                                    sb.AppendChar corZeroChar
                                End If
                                
                                sb.Append YearMod100
                            Case Else
                                sb.AppendFormat "{0:d" & Count & "}", Year
                        End Select
                        
                    Case corLowerTChar
                        If Count = 1 Then
                            sb.Append Designator, 0, 1
                        Else
                            sb.AppendString Designator
                        End If
                    
                    Case corLowerZChar
                        If ZoneOffset Is Nothing Then
                            Set ZoneOffset = TimeZone.CurrentTimeZone.GetUtcOffset(dt)
                        End If
                        
                        Select Case Count
                            Case 1
                                sb.AppendString VBA.Format$(ZoneOffset.Hours, "0;-0")
                            Case 2
                                sb.AppendString VBA.Format$(ZoneOffset.Hours, "00;-00")
                            Case Else
                                sb.AppendString VBA.Format$(ZoneOffset.Hours, "00;-00") & ":" & VBA.Format$(ZoneOffset.Minutes, "00")
                        End Select
                    
                    Case corPercentChar
                        If Percented Then
                            FormatError
                        End If
                        
                        Count = 1
                    
                    Case corBackslashChar
                        Escaped = True
                        Count = 1
                    
                    Case corLowerFChar
                        Dim secfrac As Long
                        If Count > 7 Then
                            FormatError
                        End If
                        
                        secfrac = GetSecondsFraction(dt.Ticks)
                        secfrac = (secfrac \ CLng(10 ^ (7 - Count)))
                        sb.AppendFormat "{0:g" & Count & "}", secfrac
                        
                    Case Else
                        sb.AppendChar Ch
                End Select
            End If
        Else
            sb.AppendChar Ch
            Escaped = False
            Count = 1
        End If
        Percented = (Ch = corPercentChar)
        pos = pos + Count
    Loop
    
    If Quote <> 0 Then _
        Throw Cor.NewFormatException("A matching quote was not found in the format string.")
    
    CustomFormat = StringBuilderCache.GetStringAndRelease(sb)
    Exit Function
    
Catch:
    StringBuilderCache.Release sb
    Throw
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
            sb.Append Value
        Case 2
            If Value < 10 Then
                sb.AppendChar corZeroChar
            End If
            
            sb.Append Value
            
        Case 3
            sb.AppendString ThreeCountValue
        Case Else
            sb.AppendString FourCountValue
    End Select
End Sub

Private Sub Append1or2DigitNumber(ByRef sb As StringBuilder, ByVal Count As Long, ByVal Value As Long)
    If Count = 1 Then
        sb.Append Value
    Else
        If Value < 10 Then
            sb.AppendChar corZeroChar
        End If
        
        sb.Append Value
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
    Dim i   As Long
    Dim Ch  As Integer
    
    Ch = mPattern.Chars(Index)
    i = Index + 1
    Do While i < mPattern.Buffer.cElements
        If mPattern.Chars(i) <> Ch Then
            Exit Do
        End If
        
        i = i + 1
    Loop
    GetRepeatCount = i - Index
End Function

Private Function GetPattern(ByRef Format As String) As String
    Select Case Len(Format)
        Case 0
            GetPattern = mProps.ShortDatePattern & " " & mProps.LongTimePattern
        Case 1
            Select Case AscW(Format)
                Case corLowerDChar
                    GetPattern = mProps.ShortDatePattern
                Case corUpperDChar
                    GetPattern = mProps.LongDatePattern
                Case corLowerTChar
                    GetPattern = mProps.ShortTimePattern
                Case corUpperTChar
                    GetPattern = mProps.LongTimePattern
                Case corLowerFChar
                    GetPattern = mProps.LongDatePattern & " " & mProps.ShortTimePattern
                Case corUpperFChar
                    GetPattern = Me.FullDateTimePattern
                Case corLowerGChar
                    GetPattern = mProps.ShortDatePattern & " " & mProps.ShortTimePattern
                Case corUpperGChar
                    GetPattern = mProps.ShortDatePattern & " " & mProps.LongTimePattern
                Case corLowerMChar, corUpperMChar
                    GetPattern = mProps.MonthDayPattern
                Case corLowerRChar, corUpperRChar
                    GetPattern = PatternRFC1123
                Case corLowerSChar
                    GetPattern = PatternSortableDateTime
                Case corLowerUChar
                    GetPattern = PatternUniversalSortableDateTime
                Case corUpperUChar
                    GetPattern = Me.FullDateTimePattern
                Case corLowerYChar, corUpperYChar
                    GetPattern = mProps.YearMonthPattern
                Case Else
                    Throw Cor.NewFormatException("Invalid format specifier.")
            End Select
        Case Else
            GetPattern = Format
    End Select
End Function

''
' Returns a list of date and time patterns avaiable.
Private Function CombineAllDateTimePatterns() As String()
    Dim List        As New ArrayList
    Dim Patterns    As String
    Dim Chars()     As Integer
    
    Patterns = "dDfFgGmMrRstTUuYy"
    Chars = AllocChars(Patterns)
    
    Dim i As Long
    For i = 0 To Len(Patterns) - 1
        List.AddRange GetSpecificDateTimePattern(Chars(i))
    Next i
    FreeChars Chars
    
    CombineAllDateTimePatterns = List.ToArray(vbString)
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
        Case corLowerDChar
            Ret(0) = mProps.ShortDatePattern
        Case corUpperDChar
            Ret(0) = mProps.LongDatePattern
        Case corLowerTChar
            Ret = mProps.AllShortTimePatterns
        Case corUpperTChar
            Ret(0) = mProps.LongTimePattern
        Case corUpperFChar, corUpperUChar
            Ret(0) = Me.FullDateTimePattern
        Case corLowerFChar
            Ret = CreateDateTimePatterns(mProps.AllLongDatePatterns, mProps.AllShortTimePatterns)
        Case corLowerGChar
            Ret = CreateDateTimePatterns(mProps.AllShortDatePatterns, mProps.AllShortTimePatterns)
        Case corUpperGChar
            Ret = CreateDateTimePatterns(mProps.AllShortDatePatterns, mProps.AllLongTimePatterns)
        Case corUpperMChar, corLowerMChar
            Ret(0) = mProps.MonthDayPattern
        Case corUpperRChar, corLowerRChar
            Ret(0) = PatternRFC1123
        Case corLowerSChar
            Ret(0) = PatternSortableDateTime
        Case corLowerUChar
            Ret(0) = PatternUniversalSortableDateTime
        Case corLowerYChar, corUpperYChar
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
    If mProps.IsReadOnly Then
        Error.InvalidOperation InvalidOperation_ReadOnly
    End If
End Sub

Private Function ReadStringArray(ByRef Name As String, ByVal Bag As PropertyBag) As String()
    With Bag
        Dim Count As Long
        Count = .ReadProperty(Name & "_Count", 0)
        
        Dim Values() As String
        Values = CorArray.CreateInstance(vbString, Count)
        
        Dim i As Long
        For i = 0 To Count - 1
            Values(i) = .ReadProperty(Name & "_" & i)
        Next i
        
        ReadStringArray = Values
    End With
End Function

Private Sub WriteStringArray(ByRef Arr() As String, ByRef Name As String, ByVal Bag As PropertyBag)
    With Bag
        .WriteProperty Name & "_Count", UBound(Arr) + 1
        
        Dim i As Long
        For i = 0 To UBound(Arr)
            .WriteProperty Name & "_" & i, Arr(i)
        Next i
    End With
End Sub

Private Sub VerifyLoaded()
    If Not mLoaded Then
        Init INVARIANT_LCID, False
    End If
End Sub

Private Function CopyArray(ByRef Source() As String, ByVal RequiredLength As Long) As String()
    VerifyWritable
    VerifyLoaded
    CheckValidSingleDimArray Source, Parameter_Value
    
    If CorArray.Length(Source) <> RequiredLength Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidArrayLength, RequiredLength), "Value")
    
    Dim Target() As String
    ReDim Target(0 To RequiredLength - 1)
    CorArray.CopyEx Source, LBound(Source), Target, 0, RequiredLength
    CopyArray = Target
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With mProps
        Set .Calendar = PropBag.ReadProperty(PropCalendar, Nothing)
        .CalendarWeekRule = PropBag.ReadProperty(PropCalendarWeekRule)
        .AMDesignator = PropBag.ReadProperty(PropAMDesignator)
        .DateSeparator = PropBag.ReadProperty(PropDateSeparator)
        .FirstDayOfWeek = PropBag.ReadProperty(PropFirstDayOfWeek)
        .FullDateTimePattern = PropBag.ReadProperty(PropFullDateTimePattern)
        .LongDatePattern = PropBag.ReadProperty(PropLongDatePattern)
        .LongTimePattern = PropBag.ReadProperty(PropLongTimePattern)
        .MonthDayPattern = PropBag.ReadProperty(PropMonthDayPattern)
        .PMDesignator = PropBag.ReadProperty(PropPMDesignator)
        .ShortDatePattern = PropBag.ReadProperty(PropShortDatePattern)
        .ShortTimePattern = PropBag.ReadProperty(PropShortTimePattern)
        .TimeSeparator = PropBag.ReadProperty(PropTimeSeparator)
        .YearMonthPattern = PropBag.ReadProperty(PropYearMonthPattern)
        .IsReadOnly = PropBag.ReadProperty(PropIsReadOnly)
        .AbbreviatedDayNames = ReadStringArray(PropAbbreviatedDayNames, PropBag)
        .AbbreviatedMonthNames = ReadStringArray(PropAbbreviatedMonthNames, PropBag)
        .MonthNames = ReadStringArray(PropMonthNames, PropBag)
        .DayNames = ReadStringArray(PropDayNames, PropBag)
    End With
    
    mLoaded = True
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With mProps
        PropBag.WriteProperty PropCalendar, .Calendar
        PropBag.WriteProperty PropCalendarWeekRule, .CalendarWeekRule
        WriteStringArray .AbbreviatedDayNames, PropAbbreviatedDayNames, PropBag
        WriteStringArray .AbbreviatedMonthNames, PropAbbreviatedMonthNames, PropBag
        PropBag.WriteProperty PropAMDesignator, .AMDesignator
        PropBag.WriteProperty PropDateSeparator, .DateSeparator
        WriteStringArray .DayNames, PropDayNames, PropBag
        PropBag.WriteProperty PropFirstDayOfWeek, .FirstDayOfWeek
        PropBag.WriteProperty PropFullDateTimePattern, .FullDateTimePattern
        PropBag.WriteProperty PropLongDatePattern, .LongDatePattern
        PropBag.WriteProperty PropLongTimePattern, .LongTimePattern
        PropBag.WriteProperty PropMonthDayPattern, .MonthDayPattern
        WriteStringArray .MonthNames, PropMonthNames, PropBag
        PropBag.WriteProperty PropPMDesignator, .PMDesignator
        PropBag.WriteProperty PropShortDatePattern, .ShortDatePattern
        PropBag.WriteProperty PropShortTimePattern, .ShortTimePattern
        PropBag.WriteProperty PropTimeSeparator, .TimeSeparator
        PropBag.WriteProperty PropYearMonthPattern, .YearMonthPattern
        PropBag.WriteProperty PropIsReadOnly, .IsReadOnly
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: CorDateTime
'

''
' Represents a point in time.
'
' @remarks The CorDateTime class can represent a point in time from
' 1/1/0001 12:00:00 AM to 12/31/9999 11:59:59 PM. The time is represented
' in a gregorian calendar set.
'
' <pre>
' Dim dt As CorDateTime
' Set dt = NewDate(#1/1/2005 8:30:00AM#")
' </pre>
'
' @see Constructors
' @see CorDateTimeStatic
' @see DateTimeFormatInfo
' @see TimeSpan
' @see IComparable
' @see IFormattable
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const PropMilliseconds      As String = "Milliseconds"
Private Const PropKind              As String = "Kind"

Private Const DefaultMilliseconds   As Long = 0
Private Const DefaultKind           As Long = 0

Private Const DaysPer100Years       As Long = DaysPer4Years * 25 - 1
Private Const DaysPer400Years       As Long = DaysPer100Years * 4 + 1
Private Const DaysTo10000           As Currency = DaysPer400Years * 25 - 366
Private Const MaxMilliseconds       As Currency = DaysTo10000 * MilliSecondsPerDay
Private Const DaysTo1601            As Currency = DaysPer400Years * 4
Private Const MillisecondsTo1601    As Currency = DaysTo1601 * MilliSecondsPerDay

Private Const MinFiletime           As Currency = 50491123200000@


' Currency works perfectly as the 64bit Integer in .NET. Even though the
' time is represented in milliseconds here, the fractional portion is exactly
' 10000 ticks, or .0001 is exactly 100 nanoseconds, just as in .NET. We get
' the precision as used in .NET and the 64bit Integer.
Private mMilliseconds   As Currency
Private mKind           As DateTimeKind


''
' Returns the <a href="CorDateTimeStatic.DateTimeKind.html">DateTimeKind</a> of <a href="CorDateTime.html">CorDateTime</a>.
'
' @return The <a href="CorDateTimeStatic.DateTimeKind.html">DateTimeKind</a>.
' @remarks A <a href="CorDateTime.html">CorDateTime</a> can be Universal Coordinated Time (UTC), local time, or unspecified.
'
Public Property Get Kind() As DateTimeKind
    Kind = mKind
End Property

''
' Returns if this <a href="CorDateTime.html">CorDateTime</a> instance is currently in daylight savings time.
'
' @return If this instance is in daylight savings.
' @remarks This uses the current time zone settings of the system to determine when daylight savings should be.
'
Public Property Get IsDayLightSavingsTime() As Boolean
    IsDayLightSavingsTime = TimeZone.CurrentTimeZone.IsDayLightSavingTime(Me)
End Property

''
' Converts the current instance to an Ole Automation date representation
' that is normally used in Visual Basic.
'
' @return The Visual Basic compatible date.
' @remarks It is possible that this instance holds a date outside of the
' valid range of dates that Visual Basic can represent. If this happens,
' then a normal VB error will be raised during the assignment.
'
Public Function ToOADate() As Date
    Const DaysTo1899 As Long = 693593
    
    If mMilliseconds = 0@ Then Exit Function
    
    If mMilliseconds < MilliSecondsPerDay Then
        ToOADate = 1# * mMilliseconds / MilliSecondsPerDay
    Else
        Dim Days As Double
        Days = Int(mMilliseconds / MilliSecondsPerDay)
        
        Dim Time As Double
        Time = (mMilliseconds - (Days * MilliSecondsPerDay)) / MilliSecondsPerDay
        
        Days = Days - DaysTo1899
        If Days < 0# Then
            ToOADate = Days - Time
        Else
            ToOADate = Days + Time
        End If
    End If
End Function

''
' Returns the Year part of the current date representation.
'
' @return The year of the date.
'
Public Property Get Year() As Long
    GetDateParts YearPart, Year
End Property

''
' Returns the Day part of the current date representation.
'
' @return The day of the date.
'
Public Property Get Day() As Long
    GetDateParts DayPart, Day:=Day
End Property

''
' Returns the Month part of the current date representation.
'
' @return The month of the date.
'
Public Property Get Month() As Long
    GetDateParts MonthPart, Month:=Month
End Property

''
' Returns the day within the calendar year.
'
' @return The day within the year.
' @remarks The value is calculated using the Gregorian calendar, and will range from 1-365 or 1-366 for leap years.
'
Public Property Get DayOfYear() As Long
    GetDateParts DayOfTheYear, DayOfYear:=DayOfYear
End Property

''
' Returns the day of the week for the current date.
'
' @return The day of the week.
' @remarks The <a href="CorDateTimeStatic.DayOfWeek.html">DayOfWeek</a> is zero based and therefore not
' directly compatible with vbDayOfWeek. 1 must be added to
' align to the 1-based values of vbDayOfWeek.
'
Public Property Get DayOfWeek() As DayOfWeek
    DayOfWeek = Int(mMilliseconds / MilliSecondsPerDay + 1) Mod DaysPerWeek
End Property

''
' Returns the Hour part of the current date representation.
'
' @return The current hour.
'
Public Property Get Hour() As Long
    Hour = Int(mMilliseconds / MillisecondsPerHour) Mod HoursPerDay
End Property

''
' Returns the Minute part of the current date representation.
'
' @return The current minute.
'
Public Property Get Minute() As Long
    Minute = Modulus(Int(mMilliseconds / MillisecondsPerMinute), MinutesPerHour)
End Property

''
' Returns the Second part of the current date representation.
'
' @return The current second.
'
Public Property Get Second() As Long
    Second = Modulus(Int(mMilliseconds / MillisecondsPerSecond), SecondsPerMinute)
End Property

''
' Returns the Millisecond part of the current date representation.
'
' @return The current millisecond.
'
Public Property Get Millisecond() As Long
    Millisecond = Modulus(mMilliseconds, MillisecondsPerSecond)
End Property

''
' Returns the whole date represented in milliseconds.
'
' @return The number of milliseconds from 1/1/0001.
' @remarks The return value is Currency and can represent fractions
' of milliseconds called ticks. There are 10000 ticks per millisecond,
' so the fractional portion is the number of ticks.
'
Public Property Get TotalMilliseconds() As Currency
    TotalMilliseconds = mMilliseconds
End Property

''
' Returns the whole date represented in ticks.
'
' @return The date in ticks.
' @remarks The return value is a Decimal containing the
' the number of whole ticks.
'
Public Property Get Ticks() As Variant
    Ticks = CDec(mMilliseconds) * 10000
End Property

''
' Returns the date portion of the current date representation.
'
' @return The date portion.
' @remarks The date portion does not include the time portion from the
' original date representation. The new date will have a time of 12:00AM.
'
Public Property Get DateOnly() As CorDateTime
    Set DateOnly = CorDateTime.FromMilliseconds(mMilliseconds - Modulus(mMilliseconds, MilliSecondsPerDay))
End Property

''
' Returns the time portion of the current date representation.
'
' @return The time portion.
' @remarks The time portion does not include the original date portion.
' It now will have a date of 1/1/0001.
'
Public Property Get TimeOfDay() As TimeSpan
    Set TimeOfDay = TimeSpan.FromMilliseconds(mMilliseconds - (mMilliseconds - Modulus(mMilliseconds, MilliSecondsPerDay)))
End Property

''
' Adds a specified amount of time to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param Value The amount of time to add.
' @return The new <a href="CorDateTime.html">CorDateTime</a> object containing the new calculated time.
'
Public Function Add(ByRef Value As TimeSpan) As CorDateTime
    If Value Is Nothing Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(Parameter_Value), Environment.GetResourceString(ArgumentNull_Generic))

    Set Add = CorDateTime.FromMilliseconds(mMilliseconds + Value.TotalMilliseconds, mKind)
End Function

''
' Adds a specified number of milliseconds to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of milliseconds to add.
' @return The new <a href="CorDateTime.html">CorDateTime</a> containing the new calculated time.
'
Public Function AddMilliseconds(ByVal Value As Double) As CorDateTime
    If Value < 0# Then
        Value = Fix(Value - 0.5)
    Else
        Value = Fix(Value + 0.5)
    End If
    
    If Value < -MaxMilliseconds Or Value > MaxMilliseconds Then _
        Throw New ArgumentOutOfRangeException
    
    Set AddMilliseconds = CorDateTime.FromMilliseconds(mMilliseconds + CCur(Value), mKind)
End Function

''
' Adds a specified number of days to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of days to add.
' @return The new <a href="CorDateTime.html">CorDateTime</a> containing the calculated value.
'
Public Function AddDays(ByVal Value As Double) As CorDateTime
    Set AddDays = AddMilliseconds(Value * MilliSecondsPerDay)
End Function

''
' Adds a specified number of hours to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of hours to add.
' @return The new <a href="CorDateTime.html">CorDateTime</a> instance containing the calculated value.
'
Public Function AddHours(ByVal Value As Double) As CorDateTime
    Set AddHours = AddMilliseconds(Value * MillisecondsPerHour)
End Function

''
' Adds a specified number of minutes to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of minutes to add.
' @return The new <a href="CorDateTime.html">CorDateTime</a> instance containing the calculated value.
'
Public Function AddMinutes(ByVal Value As Double) As CorDateTime
    Set AddMinutes = AddMilliseconds(Value * MillisecondsPerMinute)
End Function

''
' Adds a specified number of months to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param Months The number of months to be added.
' @return The new <a href="CorDateTime.html">CorDateTime</a> instance containing the calculated value.
' @remarks If the new month does not have as many days in it as the
' original month, then the day is set to the last day of the new month
' if it exceeds the maximum number of days for the new month.
'
Public Function AddMonths(ByVal Months As Long) As CorDateTime
    If Months < -120000 Or Months > 120000 Then _
        Throw Cor.NewArgumentOutOfRangeException("Months", "Months must be between -120000 and 120000 inclusively.")
    
    If Months = 0 Then
        Set AddMonths = CorDateTime.FromMilliseconds(mMilliseconds, mKind)
        Exit Function
    End If
    
    Dim Year    As Long
    Dim Month   As Long
    Dim Day     As Long
    GetDateParts Complete, Year, Month, Day
    
    Dim YearsToAdd  As Long
    Dim MonthsToAdd As Long
    
    YearsToAdd = CorMath.DivRem(Months, MonthsPerYear, MonthsToAdd)
    
    Month = Month + MonthsToAdd
    Year = Year + YearsToAdd
    Select Case Month
        Case Is < 1
            Month = Month + MonthsPerYear
            Year = Year - 1
        Case Is > MonthsPerYear
            Month = Month - MonthsPerYear
            Year = Year + 1
    End Select
    
    Dim MaxDays As Long
    MaxDays = CorDateTime.DaysInMonth(Year, Month)
    If Day > MaxDays Then
        Day = MaxDays
    End If
    
    Set AddMonths = CorDateTime.FromMilliseconds(CorDateTime.GetTotalDays(Year, Month, Day) * MilliSecondsPerDay + Modulus(mMilliseconds, MilliSecondsPerDay), mKind)
End Function

''
' Adds a specified number of seconds to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of seconds to be added.
' @return The new <a href="CorDateTime.html">CorDateTime</a> instance containing the calculated value.
'
Public Function AddSeconds(ByVal Value As Double) As CorDateTime
    Set AddSeconds = AddMilliseconds(Value * MillisecondsPerSecond)
End Function

''
' Adds a specified number of ticks to the current date, returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of ticks to be added.
' @return the new <a href="CorDateTime.html">CorDateTime</a> instance containing the calcuated value.
'
Public Function AddTicks(ByRef Value As Variant) As CorDateTime
    Dim DecimalValue As Variant
    DecimalValue = CDec(Value)
    
    If DecimalValue < CDec(0) Or DecimalValue > CDec(MaxMilliseconds) * 10000 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", "Ticks is outside the valid range of values")
    
    Set AddTicks = CorDateTime.FromMilliseconds(mMilliseconds + DecimalValue / 10000, mKind)
End Function

''
' Adds a specified number of years to the current date., returning a new <a href="CorDateTime.html">CorDateTime</a> instance with the calcuated value.
'
' @param value The number of years to be added.
' @return The new <a href="CorDateTime.html">CorDateTime</a> instance containing the calculated value.
'
Public Function AddYears(ByVal Value As Long) As CorDateTime
    Set AddYears = AddMonths(Value * MonthsPerYear)
End Function

''
' Compares this instance to a passed in Date or <a href="CorDateTime.html">CorDateTime</a> instance.
'
' @param Value The Date or <a href="CorDateTime.html">CorDateTime</a> value to compare against this instance.
' @return A value indicating the relation to the passed in value. Negative
' indicates this instance is less than the argument, zero indicates the two
' values are equal, and positive indicates this value is greater than the
' argument.
'
Public Function CompareTo(ByRef Value As Variant) As Long
    If VarType(Value) = vbObject Then
        If Value Is Nothing Then
            CompareTo = 1
            Exit Function
        End If
    End If
    
    Select Case Statics.CorDateTime.GetTotalMilliseconds(Value)
        Case Is < mMilliseconds
            CompareTo = 1
        Case Is > mMilliseconds
            CompareTo = -1
    End Select
End Function

''
' Converts the <a href="CorDateTime.html">CorDateTime</a> to a FileTime representation, adjusting for timezone.
'
' @return The filetime representation of the <a href="CorDateTime.html">CorDateTime</a>.
' @remarks FileTime values cannot be less than 1/1/1601 12:00:00AM.
'
Public Function ToFileTime() As Currency
    ToFileTime = ToUniversalTime.ToFileTimeUtc
End Function

''
' Converts the <a href="CorDateTime.html">CorDateTime</a> to a FileTime representation, ignoring the timezone offset.
'
' @return The filetime representation of the <a href="CorDateTime.html">CorDateTime</a>.
' @remarks FileTime values cannot be less than 1/1/1601 12:00:00AM.
'
Public Function ToFileTimeUtc() As Currency
    Dim ms As Currency
    
    If mKind = LocalKind Then
        ms = ToUniversalTime.TotalMilliseconds
    Else
        ms = mMilliseconds
    End If
    
    ms = ms - MillisecondsTo1601
    
    If ms < 0 Then _
        Error.ArgumentOutOfRange vbNullString, ArgumentOutOfRange_InvalidFileTime
    
    ToFileTimeUtc = ms
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> with the timezone offset applied to this instance.
'
' @return The <a href="CorDateTime.html">CorDateTime</a> object with the timezone offset applied.
'
Public Function ToLocalTime() As CorDateTime
    If mKind = LocalKind Then
        Set ToLocalTime = Me
    Else
        Set ToLocalTime = TimeZone.CurrentTimeZone.ToLocalTime(Me)
    End If
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> with the timezone offset removed from this instance.
'
' @return The <a href="CorDateTime.html">CorDateTime</a> object with the timezone offset removed.
'
Public Function ToUniversalTime() As CorDateTime
    If mKind = UtcKind Then
        Set ToUniversalTime = Me
    Else
        Set ToUniversalTime = TimeZone.CurrentTimeZone.ToUniversalTime(Me)
    End If
End Function

''
' Formats a <a href="CorDateTime.html">CorDateTime</a> object or Date value in the pattern specified by
' the formatting command. If no command is specified, then 'G' is assumed.
' If the command cannot be found, then a custom pattern is assumed.
'
' @param Format The format inwhich the date should represent.
' @param provider A formatting provider to format the date.
' @return String representing this instance.
' @see DateTimeFormatInfo
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
Public Function ToString(Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim Info As DateTimeFormatInfo
    
    Set Info = DateTimeFormatInfo.GetInstance(Provider)
    ToString = Info.Format(Me, Format)
End Function

''
' Returns the date formatted to the current cultures long date pattern.
'
' @return Long date pattern formatted string.
'
Public Function ToLongDateString() As String
    ToLongDateString = DateTimeFormatInfo.CurrentInfo.Format(Me, "D")
End Function

''
' Returns the date formatted to the current cultures short date pattern.
'
' @return Short date pattern formatted string.
'
Public Function ToShortDateString() As String
    ToShortDateString = DateTimeFormatInfo.CurrentInfo.Format(Me, "d")
End Function

''
' Returns the time formatted to the current cultures long time pattern.
'
' @return Long time pattern formatted string.
'
Public Function ToLongTimeString() As String
    ToLongTimeString = DateTimeFormatInfo.CurrentInfo.Format(Me, "T")
End Function

''
' Returns the time formatted to the current cultures short time pattern.
'
' @return Short time pattern formatted string.
'
Public Function ToShortTimeString() As String
    ToShortTimeString = DateTimeFormatInfo.CurrentInfo.Format(Me, "t")
End Function

''
' Returns a 64-bit representation of the time and kind.
'
' @return 64-bit representation of the <a href="CorDateTime.html">CorDateTime</a>
' @remarks The upper 2 bits are used to represent the DateTimeKind
' of the object, the lower 62-bits are used to represent the ticks.
'
Public Function ToBinary() As Currency
    ToBinary = mMilliseconds
    
    If mKind = LocalKind Then
        Dim ts As TimeSpan
        Set ts = TimeZone.CurrentTimeZone.GetUtcOffset(Me)
        ToBinary = ToBinary - ts.TotalMilliseconds
    End If
    
    MemByte(VarPtr(ToBinary) + 7) = MemByte(VarPtr(ToBinary) + 7) Or (mKind * &H40)
End Function

''
' Returns a boolean indicating if the value and this object instance are the same date.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    On Error GoTo Catch
    Equals = Statics.CorDateTime.GetTotalMilliseconds(Value) = mMilliseconds
Catch:
End Function

''
' Subtracts either a <a href="CorDateTime.html">CorDateTime</a> instance or a <a href="TimeSpan.html">TimeSpan</a> instance from this
' CorDateTime instance.
'
' @param value A <a href="CorDateTime.html">CorDateTime</a> or <a href="TimeSpan.html">TimeSpan</a> object to subtract from this instance.
' @return If subtracting a <a href="CorDateTime.html">CorDateTime</a> object from this, then a <a href="TimeSpan.html">TimeSpan</a>
' representing the difference in time between the two is returned. If subtracting a <a href="TimeSpan.html">TimeSpan</a> object from this instance, then a new <a href="CorDateTime.html">CorDateTime</a> with a difference
' of the <a href="TimeSpan.html">TimeSpan</a> obect from this instance is returned.
'
Public Function Subtract(ByRef Value As Variant) As Object
    Select Case VarType(Value)
        Case vbObject
            If Value Is Nothing Then _
                Error.ArgumentNull "Value"
            
            If TypeOf Value Is TimeSpan Then
                Dim ts As TimeSpan
                Set ts = Value
                Set Subtract = CorDateTime.FromMilliseconds(mMilliseconds - ts.TotalMilliseconds, mKind)
            ElseIf TypeOf Value Is CorDateTime Then
                Dim dt As CorDateTime
                Set dt = Value
                Set Subtract = TimeSpan.FromMilliseconds(mMilliseconds - dt.TotalMilliseconds)
            Else
                Error.Argument Argument_InvalidDateSubtraction, "Value"
            End If
        Case vbDate
            Set Subtract = TimeSpan.FromMilliseconds(mMilliseconds - DateToMilliseconds(Value))
        Case Else
            Error.Argument Argument_InvalidDateSubtraction, "Value"
    End Select
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    With AsDLong(mMilliseconds)
        GetHashCode = .HiDWord Xor .LoDWord
    End With
End Function

''
' Determines if this instance is greater than another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is greater than the <i>Value</i> object.
' @remarks This is provided for quick comparisions of two <a href="CorDateTime.html">CorDateTime</a> objects.
'
Public Function GreaterThan(ByVal Value As CorDateTime) As Boolean
    If Value Is Nothing Then
        GreaterThan = True
    Else
        GreaterThan = (mMilliseconds > Value.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is less than another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is less than the <i>Value</i> object.
' @remarks This is provided for quick comprisons of two <a href="CorDateTime.html">CorDateTime</a> Objects.
'
Public Function LessThan(ByVal Value As CorDateTime) As Boolean
    If Not Value Is Nothing Then
        LessThan = (mMilliseconds < Value.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is greater than or equal to another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is greater than or equal to the <i>Value</i> object.
' @remarks This is provided for quick comprisons of two <a href="CorDateTime.html">CorDateTime</a> Objects.
'
Public Function GreaterThanOrEqualTo(ByVal Value As CorDateTime) As Boolean
    If Value Is Nothing Then
        GreaterThanOrEqualTo = True
    Else
        GreaterThanOrEqualTo = (mMilliseconds >= Value.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is less than or equal to another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is less than or equal to the <i>Value</i> object.
' @remarks This is provided for quick comprisons of two <a href="CorDateTime.html">CorDateTime</a> Objects.
'
Public Function LessThanOrEqualTo(ByVal Value As CorDateTime) As Boolean
    If Not Value Is Nothing Then
        LessThanOrEqualTo = (mMilliseconds <= Value.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is equal to another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is equal to the <i>Value</i> object.
' @remarks This is provided for quick comprisons of two <a href="CorDateTime.html">CorDateTime</a> Objects.
'
Public Function EqualTo(ByVal Value As CorDateTime) As Boolean
    If Not Value Is Nothing Then
        EqualTo = (mMilliseconds = Value.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is not equal to another <a href="CorDateTime.html">CorDateTime</a> object.
'
' @param Value The <a href="CorDateTime.html">CorDateTime</a> object to compare against.
' @return Returns True if this object is not equal to the <i>Value</i> object.
' @remarks This is provided for quick comprisons of two <a href="CorDateTime.html">CorDateTime</a> Objects.
'
Public Function NotEqualTo(ByVal Value As CorDateTime) As Boolean
    If Value Is Nothing Then
        NotEqualTo = True
    Else
        NotEqualTo = (mMilliseconds <> Value.TotalMilliseconds)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, ByVal Calendar As Calendar, ByVal Kind As DateTimeKind)
    If Calendar Is Nothing Then
        mMilliseconds = CorDateTime.GetTotalDays(Year, Month, Day) * MilliSecondsPerDay + GetTotalMilliseconds(Hour, Minute, Second, Millisecond)
    Else
        mMilliseconds = Calendar.ToDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond).TotalMilliseconds
    End If
    mKind = Kind
End Sub

Friend Sub InitFromMilliseconds(ByVal ms As Currency, ByVal Kind As DateTimeKind)
    If ms < 0@ Or ms > MaxMilliseconds Then _
        Throw Cor.NewArgumentOutOfRangeException("ms", "Milliseconds must be between 0 and " & MaxMilliseconds & " inclusively.")
        
    mMilliseconds = ms
    mKind = Kind
End Sub

Friend Sub InitFromDate(ByVal d As Date, ByVal Kind As DateTimeKind)
    mMilliseconds = DateToMilliseconds(d)
    mKind = Kind
End Sub

Friend Sub InitFromCorDateTime(ByVal d As CorDateTime, ByVal Kind As DateTimeKind)
    If Not d Is Nothing Then
        mMilliseconds = d.TotalMilliseconds
    End If
    
    mKind = Kind
End Sub

Friend Sub InitFromDateOrDateTime(ByRef Time As Variant, ByVal Kind As DateTimeKind)
    Select Case VarType(Time)
        Case vbObject
            InitFromCorDateTime Time, Kind
        Case vbDate
            InitFromDate Time, Kind
        Case Else
            Error.Argument Arg_MustBeDateTime
    End Select
End Sub

Friend Sub InitFromFileTime(ByVal Time As Currency, ByVal Kind As DateTimeKind)
    mMilliseconds = Time + MillisecondsTo1601
    mKind = Kind
End Sub

Friend Sub InitFromBinary(ByVal Value As Currency)
    mMilliseconds = Value

    ' Get the address of the 8th byte in the 64-bit value.
    Dim PtrToHighByte As Long
    PtrToHighByte = VarPtr(mMilliseconds) + 7
    
    ' Get the value of the 8th byte in the 64-bit value.
    Dim HighByte As Long
    HighByte = MemByte(PtrToHighByte)
    
    ' Strip off the highest 2-bits of the 8th byte in the 64-bit value.
    MemByte(PtrToHighByte) = HighByte And &H3F
    
    ' Shift the upper 2 bits to the lowest 2 bits to represent the Kind.
    mKind = (HighByte And &HC0) \ &H40
    
    If mKind = LocalKind Then
        Dim ts As TimeSpan
        Set ts = TimeZone.CurrentTimeZone.GetUtcOffset(Me)
        mMilliseconds = mMilliseconds + ts.TotalMilliseconds
    End If
End Sub

Friend Sub GetDateParts(ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim Days As Long
    Days = Int(Int(mMilliseconds) / MilliSecondsPerDay)
    
    Dim Years400 As Long
    Years400 = Days \ DaysPer400Years
    Days = Days - Years400 * DaysPer400Years
    
    Dim Years100 As Long
    Years100 = Days \ DaysPer100Years
    If Years100 = 4 Then Years100 = 3   ' leap year
    Days = Days - Years100 * DaysPer100Years
    
    Dim Years4 As Long
    Years4 = Days \ DaysPer4Years
    Days = Days - Years4 * DaysPer4Years
    
    Dim Years As Long
    Years = Days \ 365
    If Years = 4 Then Years = 3 ' leap year
    
    ' Return the Year.
    Year = Years400 * 400 + Years100 * 100 + Years4 * 4 + Years + 1
    If DatePart = YearPart Then Exit Sub
    
    Days = Days - Years * DaysPerYear
    
    ' Return the DayOfYear
    DayOfYear = Days + 1
    If DatePart = DayOfTheYear Then Exit Sub
    
    ' Month will be set in the CalculateDay method.
    ' Day is set to the result of CalculateDays.
    If Years = 3 And (Years100 = 3 Or Years4 <> 24) Then
        Day = CalculateDay(Days, Month, DaysToMonthLeapYear) + 1
    Else
        Day = CalculateDay(Days, Month, DaysToMonth) + 1
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalculateDay(ByVal Days As Long, ByRef Month As Long, ByRef Months() As Long) As Long
    Month = 0
    
    Do While Months(Month) <= Days
        Month = Month + 1
    Loop
    
    CalculateDay = Days - Months(Month - 1)
End Function

Private Function GetTotalMilliseconds(ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long) As Currency
    GetTotalMilliseconds = Hour * MillisecondsPerHour + Minute * MillisecondsPerMinute + Second * MillisecondsPerSecond + Millisecond
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMilliseconds = .ReadProperty(PropMilliseconds, DefaultMilliseconds)
        mKind = .ReadProperty(PropKind, DefaultKind)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropMilliseconds, mMilliseconds
        .WriteProperty PropKind, mKind
    End With
End Sub


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

Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByRef Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDateTime"
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
'    Module: cDateTime
'

''
' Represents a point in time.
'
' @remarks The cDateTime class can represent a point in time from
' 1/1/0001 12:00:00 AM to 12/31/9999 11:59:59 PM. The time is represented
' in a gregorian calendar set.
'
' <pre>
' Dim dt As cDateTime
' Set dt = NewDate(#1/1/2005 8:30:00AM#")
' </pre>
'
' @see Constructors
' @see cDateTimeStatic
' @see DateTimeFormatInfo
' @see TimeSpan
' @see IComparable
' @see IFormattable
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const PROP_MILLISECONDS         As String = "Milliseconds"
Private Const PROP_KIND                 As String = "Kind"
Private Const DEF_MILLISECONDS          As Long = 0@
Private Const DEF_KIND                  As Long = 0

Private Const SECONDS_PER_MINUTE        As Currency = 60@
Private Const MINUTES_PER_HOUR          As Currency = 60@
Private Const HOURS_PER_DAY             As Long = 24
Private Const MONTHS_PER_YEAR           As Long = 12

Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private Const DAYS_TO_18991231          As Long = 693593
Private Const MILLISECONDS_TO_18991231  As Currency = 1@ * DAYS_TO_18991231 * MILLISECONDS_PER_DAY
Private Const DAYS_PER_YEAR             As Long = 365
Private Const DAYS_PER_4_YEARS          As Long = DAYS_PER_YEAR * 4 + 1
Private Const DAYS_PER_100_YEARS        As Long = DAYS_PER_4_YEARS * 25 - 1
Private Const DAYS_PER_400_YEARS        As Long = DAYS_PER_100_YEARS * 4 + 1
Private Const DAYS_TO_YEAR_10000        As Currency = DAYS_PER_400_YEARS * 25 - 366
Private Const MAX_MILLISECONDS          As Currency = DAYS_TO_YEAR_10000 * MILLISECONDS_PER_DAY
Private Const DAYS_TO_16011231          As Currency = DAYS_PER_400_YEARS * 4
Private Const MILLISECONDS_TO_16011231  As Currency = DAYS_TO_16011231 * MILLISECONDS_PER_DAY

Private Const FILETIME_MINIMUM          As Currency = 50491123200000@

' Currency works perfectly as the 64bit Integer in .NET. Even though the
' time is represented in milliseconds here, the fractional portion is exactly
' 10000 ticks, or .0001 is exactly 100 nanoseconds, just as in .NET. We get
' the precision as used in .NET and the 64bit Integer.
Private mMilliseconds   As Currency
Private mKind           As DateTimeKind


''
' Returns the kind of cDateTime.
'
' @return The kind.
' @remarks A <b>cDateTime</b> can be Universal Coordinated Time (UTC),
' local time, or unspecified.
'
Public Property Get Kind() As DateTimeKind
    Kind = mKind
End Property

''
' Returns if this cDateTime instance is currently in daylight savings time.
'
' @return If this instance is in daylight savings.
' @remarks This uses the current time zone settings of the system to
' determine when daylight savings should be.
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
    If mMilliseconds = 0@ Then Exit Function
    
    If mMilliseconds < MILLISECONDS_PER_DAY Then
        ToOADate = 1# * mMilliseconds / MILLISECONDS_PER_DAY
    Else
        Dim Days As Double
        Days = Int(mMilliseconds / MILLISECONDS_PER_DAY)
        
        Dim Time As Double
        Time = (mMilliseconds - (Days * MILLISECONDS_PER_DAY)) / MILLISECONDS_PER_DAY
        
        Days = Days - DAYS_TO_18991231
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
    Call GetDateParts(YearPart, Year)
End Property

''
' Returns the Day part of the current date representation.
'
' @return The day of the date.
'
Public Property Get Day() As Long
    Call GetDateParts(DayPart, , , Day)
End Property

''
' Returns the Month part of the current date representation.
'
' @return The month of the date.
'
Public Property Get Month() As Long
    Call GetDateParts(MonthPart, , Month)
End Property

''
' Returns the day within the calendar year.
'
' @return The day within the year.
' @remarks The value is calculated using the Gregorian calendar,
' and will range from 1-365 or 1-366 for leap years.
'
Public Property Get DayOfYear() As Long
    Call GetDateParts(DayOfTheYear, , , , DayOfYear)
End Property

''
' Returns the day of the week for the current date.
'
' @return The day of the week.
' @remarks The DayOfWeekEnum is zero based and therefore not
' directly compatible with vbDayOfWeek. 1 must be added to
' align to the 1-based values of vbDayOfWeek.
'
Public Property Get DayOfWeek() As DayOfWeek
    DayOfWeek = Int(mMilliseconds / MILLISECONDS_PER_DAY + 1) Mod 7
End Property

''
' Returns the Hour part of the current date representation.
'
' @return The current hour.
'
Public Property Get Hour() As Long
    Hour = Int(mMilliseconds / MILLISECONDS_PER_HOUR) Mod HOURS_PER_DAY
End Property

''
' Returns the Minute part of the current date representation.
'
' @return The current minute.
'
Public Property Get Minute() As Long
    'Minute = Int(mMilliseconds / MILLISECONDS_PER_MINUTE) Mod MINUTES_PER_HOUR
    Minute = Modulus(Int(mMilliseconds / MILLISECONDS_PER_MINUTE), MINUTES_PER_HOUR)
End Property

''
' Returns the Second part of the current date representation.
'
' @return The current second.
'
Public Property Get Second() As Long
    Second = Modulus(Int(mMilliseconds / MILLISECONDS_PER_SECOND), SECONDS_PER_MINUTE)
End Property

''
' Returns the Millisecond part of the current date representation.
'
' @return The current millisecond.
'
Public Property Get Millisecond() As Long
    Millisecond = Modulus(mMilliseconds, MILLISECONDS_PER_SECOND)
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
Public Property Get DateOnly() As cDateTime
    Set DateOnly = cDateTime.FromMilliseconds(mMilliseconds - Modulus(mMilliseconds, MILLISECONDS_PER_DAY))
End Property

''
' Returns the time portion of the current date representation.
'
' @return The time portion.
' @remarks The time portion does not include the original date portion.
' It now will have a date of 1/1/0001.
'
Public Property Get TimeOfDay() As TimeSpan
    Set TimeOfDay = TimeSpan.FromMilliseconds(mMilliseconds - (mMilliseconds - Modulus(mMilliseconds, MILLISECONDS_PER_DAY)))
End Property

''
' Adds a specified amount of time to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param ts The amount of time to add.
' @return The new date object containing the new calculated time.
'
Public Function Add(ByRef ts As TimeSpan) As cDateTime
    If ts Is Nothing Then _
        Throw Cor.NewArgumentNullException("ts", Environment.GetResourceString(ArgumentNull_TimeSpan, "ts"))
    
    Set Add = cDateTime.FromMilliseconds(mMilliseconds + ts.TotalMilliseconds, mKind)
End Function

''
' Adds a specified number of milliseconds to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of milliseconds to add.
' @return The new date containing the new calculated time.
'
Public Function AddMilliseconds(ByVal Value As Double) As cDateTime
    If Value < 0# Then
        Value = Fix(Value - 0.5)
    Else
        Value = Fix(Value + 0.5)
    End If
    
    If Value < -MAX_MILLISECONDS Or Value > MAX_MILLISECONDS Then _
        Throw New ArgumentOutOfRangeException
    
    Set AddMilliseconds = cDateTime.FromMilliseconds(mMilliseconds + CCur(Value), mKind)
End Function

''
' Adds a specified number of days to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of days to add.
' @return The new cDateTime containing the calculated value.
'
Public Function AddDays(ByVal Value As Double) As cDateTime
    Set AddDays = AddMilliseconds(Value * MILLISECONDS_PER_DAY)
End Function

''
' Adds a specified number of hours to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of hours to add.
' @return The new cDateTime instance containing the calculated value.
'
Public Function AddHours(ByVal Value As Double) As cDateTime
    Set AddHours = AddMilliseconds(Value * MILLISECONDS_PER_HOUR)
End Function

''
' Adds a specified number of minutes to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of minutes to add.
' @return The new cDateTime instance containing the calculated value.
'
Public Function AddMinutes(ByVal Value As Double) As cDateTime
    Set AddMinutes = AddMilliseconds(Value * MILLISECONDS_PER_MINUTE)
End Function

''
' Adds a specified number of months to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param Months The number of months to be added.
' @return The new cDateTime instance containing the calculated value.
' @remarks If the new month does not have as many days in it as the
' original month, then the day is set to the last day of the new month
' if it exceeds the maximum number of days for the new month.
'
Public Function AddMonths(ByVal Months As Long) As cDateTime
    If Months < -120000 Or Months > 120000 Then _
        Throw Cor.NewArgumentOutOfRangeException("Months", Months, "Months must be between -120000 and 120000 inclusively.")
    
    If Months = 0 Then
        Set AddMonths = Me
        Exit Function
    End If
    
    Dim Year    As Long
    Dim Month   As Long
    Dim Day     As Long
    Call GetDateParts(Complete, Year, Month, Day)
    
    Dim yearsToAdd As Long
    Dim monthsToAdd As Long
    
    yearsToAdd = MathExt.DivRem(Months, MONTHS_PER_YEAR, monthsToAdd)
    
    Month = Month + monthsToAdd
    Year = Year + yearsToAdd
    Select Case Month
        Case Is < 1
            Month = Month + MONTHS_PER_YEAR
            Year = Year - 1
        Case Is > MONTHS_PER_YEAR
            Month = Month - MONTHS_PER_YEAR
            Year = Year + 1
    End Select
    
    Dim MaxDays As Long
    MaxDays = cDateTime.DaysInMonth(Year, Month)
    If Day > MaxDays Then Day = MaxDays
    Set AddMonths = cDateTime.FromMilliseconds(GetTotalDays(Year, Month, Day) * MILLISECONDS_PER_DAY + Modulus(mMilliseconds, MILLISECONDS_PER_DAY), mKind)
End Function

''
' Adds a specified number of seconds to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of seconds to be added.
' @return The new cDateTime instance containing the calculated value.
'
Public Function AddSeconds(ByVal Value As Double) As cDateTime
    Set AddSeconds = AddMilliseconds(Value * MILLISECONDS_PER_SECOND)
End Function

''
' Adds a specified number of ticks to the current date, returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of ticks to be added.
' @return the new cDateTime instance containing the calcuated value.
'
Public Function AddTicks(ByRef Value As Variant) As cDateTime
    Dim DecimalValue As Variant
    DecimalValue = CDec(Value)
    
    If DecimalValue < CDec(0) Or DecimalValue > CDec(MAX_MILLISECONDS) * 10000 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value", DecimalValue, "Ticks is outside the valid range of values")
    
    Set AddTicks = cDateTime.FromMilliseconds(mMilliseconds + DecimalValue / 10000, mKind)
End Function

''
' Adds a specified number of years to the current date., returning a
' new cDateTime instance with the calcuated value.
'
' @param value The number of years to be added.
' @return The new cDateTime instance containing the calculated value.
'
Public Function AddYears(ByVal Value As Long) As cDateTime
    Set AddYears = AddMonths(Value * MONTHS_PER_YEAR)
End Function

''
' Compares this instance to a passed in cDateTime instance.
'
' @param value The cDateTime or Date value to compare against this instance.
' @return A value indicating the relation to the passed in value. Negative
' indicates this instance is less than the argument, zero indicates the two
' values are equal, and positive indicates this value is greater than the
' argument.
'
Public Function CompareTo(ByRef Value As Variant) As Long
    Dim ms As Currency
    
    Select Case VariantType(Value) And &HFF
        Case vbObject
            If Value Is Nothing Then
                CompareTo = 1
            Else
                Dim dt As cDateTime
                On Error GoTo errTrap
                Set dt = Value
                ms = dt.TotalMilliseconds
                If mMilliseconds < ms Then
                    CompareTo = -1
                ElseIf mMilliseconds > ms Then
                    CompareTo = 1
                End If
            End If
        Case vbDate
            ms = cDateTime.AscDateTime(Value).TotalMilliseconds
            If mMilliseconds < ms Then
                CompareTo = -1
            ElseIf mMilliseconds > ms Then
                CompareTo = 1
            End If
        Case vbNull, vbEmpty, vbError
            CompareTo = 1
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "value")
    End Select
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_DateRequired), "value")
End Function

''
' Converts the cDateTime to a FileTime representation, adjusting for timezone.
'
' @return The filetime representation of the cDateTime.
' @remarks FileTime values cannot be less than 1/1/1601 12:00:00AM.
'
Public Function ToFileTime() As Currency
    Dim ms As Currency
    ms = TimeZone.CurrentTimeZone.ToUniversalTime(Me).TotalMilliseconds
    If ms < FILETIME_MINIMUM Then _
        Throw Cor.NewArgumentOutOfRangeException(Message:=Resources.GetString(ArgumentOutOfRange_InvalidFileTime))
    ToFileTime = ms - MILLISECONDS_TO_16011231
End Function

''
' Converts the cDateTime to a FileTime representation, ignoring the timezone offset.
'
' @return The filetime representation of the cDateTime.
' @remarks FileTime values cannot be less than 1/1/1601 12:00:00AM.
'
Public Function ToFileTimeUtc() As Currency
    If mMilliseconds < FILETIME_MINIMUM Then _
        Throw Cor.NewArgumentOutOfRangeException(Message:=Resources.GetString(ArgumentOutOfRange_InvalidFileTime))
    ToFileTimeUtc = mMilliseconds - MILLISECONDS_TO_16011231
End Function

''
' Returns a cDateTime with the timezone offset applied to this instance.
'
' @return The cDateTime object with the timezone offset applied.
'
Public Function ToLocalTime() As cDateTime
    Set ToLocalTime = TimeZone.CurrentTimeZone.ToLocalTime(Me)
End Function

''
' Returns a cDateTime with the timezone offset removed from this instance.
'
' @return The cDateTime object with the timezone offset removed.
'
Public Function ToUniversalTime() As cDateTime
    Set ToUniversalTime = TimeZone.CurrentTimeZone.ToUniversalTime(Me)
End Function

''
' Returns a string representation of this object instance.
'
' @param Format The format inwhich the date should represent.
' @param provider A formatting provider to format the date.
' @return String representing this instance.
' @see DateTimeFormatInfo
'
Public Function ToString(Optional ByVal Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim Info As DateTimeFormatInfo
    If Not Provider Is Nothing Then Set Info = Provider.GetFormat("DateTimeFormatInfo")
    If Info Is Nothing Then Set Info = CultureInfo.CurrentCulture.DateTimeFormat
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
' @return 64-bit representation of the cDateTime
' @remarks The upper 2 bits are used to represent the DateTimeKind
' of the object, the lower 62-bits are used to represent the 100-nanosecond
' segmented time.
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
' Returns a boolean indicating if the value and this object
' instance are the same date.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbObject
            If Value Is Nothing Then Exit Function
            If TypeOf Value Is cDateTime Then
                Dim dt As cDateTime
                Set dt = Value
                Equals = (mMilliseconds = dt.TotalMilliseconds)
            End If
        Case vbDate
            Equals = (mMilliseconds = cDateTime.AscDateTime(Value).TotalMilliseconds)
    End Select
End Function

''
' Subtracts either a cDateTime instance or a TimeSpan instance from this
' cDateTime instance.
'
' @param value A cDateTime or TimeSpan object to subtract from this instance.
' @return If subtracting a cDateTime object from this, then a TimeSpan
' representing the difference in time between the two is returned. If subtracting
' a TimeSpan object from this instance, then a new cDateTime with a difference
' of the TimeSpan obect from this instance is returned.
'
Public Function Subtract(ByRef Value As Variant) As Object
    Select Case VariantType(Value) And &HFF
        Case vbObject
            If TypeOf Value Is TimeSpan Then
                Dim ts As TimeSpan
                Set ts = Value
                Set Subtract = cDateTime.FromMilliseconds(mMilliseconds - ts.TotalMilliseconds, mKind)
            ElseIf TypeOf Value Is cDateTime Then
                Dim dt As cDateTime
                Set dt = Value
                Set Subtract = TimeSpan.FromMilliseconds(mMilliseconds - dt.TotalMilliseconds)
            Else
                Throw Cor.NewArgumentException("Date, cDateTime or TimeSpan object is required.", "value")
            End If
        Case vbDate
            Set Subtract = TimeSpan.FromMilliseconds(mMilliseconds - cDateTime.AscDateTime(Value).TotalMilliseconds)
        Case Else
            Throw Cor.NewArgumentException("Date, cDateTime or TimeSpan object is required.", "value")
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
' Determines if this instance is greater than another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is greater than the dt object.
' @remarks This is provided for quick comparisions of two cDateTime objects.
'
Public Function GreaterThan(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then
        GreaterThan = True
    Else
        GreaterThan = (mMilliseconds > dt.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is less than another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is less than the dt object.
' @remarks This is provided for quick comprisons of two cDateTimeObjects.
'
Public Function LessThan(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then Exit Function
    LessThan = (mMilliseconds < dt.TotalMilliseconds)
End Function

''
' Determines if this instance is greater than or equal to another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is greater than or equal to the dt object.
' @remarks This is provided for quick comprisons of two cDateTimeObjects.
'
Public Function GreaterThanOrEqualTo(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then
        GreaterThanOrEqualTo = True
    Else
        GreaterThanOrEqualTo = (mMilliseconds >= dt.TotalMilliseconds)
    End If
End Function

''
' Determines if this instance is less than or equal to another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is less than or equal to the dt object.
' @remarks This is provided for quick comprisons of two cDateTimeObjects.
'
Public Function LessThanOrEqualTo(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then Exit Function
    LessThanOrEqualTo = (mMilliseconds <= dt.TotalMilliseconds)
End Function

''
' Determines if this instance is equal to another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is equal to the dt object.
' @remarks This is provided for quick comprisons of two cDateTimeObjects.
'
Public Function EqualTo(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then Exit Function
    EqualTo = (mMilliseconds = dt.TotalMilliseconds)
End Function

''
' Determines if this instance is not equal to another cDateTime object.
'
' @param dt The cDateTime object to compare against.
' @return Returns True if this object is not equal to the dt object.
' @remarks This is provided for quick comprisons of two cDateTimeObjects.
'
Public Function NotEqualTo(ByRef dt As cDateTime) As Boolean
    If dt Is Nothing Then
        NotEqualTo = True
    Else
        NotEqualTo = (mMilliseconds <> dt.TotalMilliseconds)
    End If
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, ByVal Calendar As Calendar, ByVal Kind As DateTimeKind)
    If Calendar Is Nothing Then
        mMilliseconds = GetTotalDays(Year, Month, Day) * MILLISECONDS_PER_DAY + GetTotalMilliseconds(Hour, Minute, Second, Millisecond)
    Else
        mMilliseconds = Calendar.ToDateTime(Year, Month, Day, Hour, Minute, Second, Millisecond).TotalMilliseconds
    End If
    mKind = Kind
End Sub

Friend Sub InitFromMilliseconds(ByVal ms As Currency, ByVal Kind As DateTimeKind)
    If ms < 0@ Or ms > MAX_MILLISECONDS Then _
        Throw Cor.NewArgumentOutOfRangeException("ms", ms, "Milliseconds must be between 0 and " & MAX_MILLISECONDS & " inclusively.")
        
    mMilliseconds = ms
    mKind = Kind
End Sub

Friend Sub InitFromDate(ByVal d As Date, ByVal Kind As DateTimeKind)
    Dim Days As Currency
    
    If d < 0# Then
        Days = Fix(d * MILLISECONDS_PER_DAY - 0.5)
        Days = Days - Modulus(Days, MILLISECONDS_PER_DAY) * 2
    Else
        Days = Fix(d * MILLISECONDS_PER_DAY + 0.5)
    End If
    mMilliseconds = Days + MILLISECONDS_TO_18991231
    mKind = Kind
End Sub

Friend Sub InitFromFileTime(ByVal Time As Currency, ByVal Kind As DateTimeKind)
    If Time < 0@ Then _
        Throw Cor.NewArgumentOutOfRangeException("Time", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    mMilliseconds = Time + MILLISECONDS_TO_16011231
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetDateParts(ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim Days As Long
    Days = Int(mMilliseconds / MILLISECONDS_PER_DAY)
    
    Dim years400 As Long
    years400 = Days \ DAYS_PER_400_YEARS
    Days = Days - years400 * DAYS_PER_400_YEARS
    
    Dim years100 As Long
    years100 = Days \ DAYS_PER_100_YEARS
    If years100 = 4 Then years100 = 3   ' leap year
    Days = Days - years100 * DAYS_PER_100_YEARS
    
    Dim years4 As Long
    years4 = Days \ DAYS_PER_4_YEARS
    Days = Days - years4 * DAYS_PER_4_YEARS
    
    Dim Years As Long
    Years = Days \ 365
    If Years = 4 Then Years = 3 ' leap year
    
    ' Return the Year.
    Year = years400 * 400 + years100 * 100 + years4 * 4 + Years + 1
    If DatePart = YearPart Then Exit Sub
    
    Days = Days - Years * DAYS_PER_YEAR
    
    ' Return the DayOfYear
    DayOfYear = Days + 1
    If DatePart = DayOfTheYear Then Exit Sub
    
    ' Month will be set in the CalculateDays method.
    ' Day is set to the result of CalculateDays.
    If Years = 3 And (years100 = 3 Or years4 <> 24) Then
        Day = CalculateDay(Days, Month, DaysToMonthLeapYear) + 1
    Else
        Day = CalculateDay(Days, Month, DaysToMonth) + 1
    End If
End Sub

Private Function CalculateDay(ByVal Days As Long, ByRef Month As Long, ByRef Months() As Long) As Long
    Do While Months(Month) <= Days
        Month = Month + 1
    Loop
    CalculateDay = Days - Months(Month - 1)
End Function

Private Function GetTotalDays(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    If Year = 0 And Month = 0 And Day = 0 Then Exit Function
    If Year < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Year", Year, Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    If Month < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Month", Month, Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    If Day < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Day, Environment.GetResourceString(ArgumentOutOfRange_ValidValues))
    
    Dim Days As Long
    If cDateTime.IsLeapYear(Year) Then
        Days = DaysToMonthLeapYear(Month - 1)
    Else
        Days = DaysToMonth(Month - 1)
    End If
    
    Year = Year - 1
    GetTotalDays = (Day - 1) + Days + 365 * Year + (Year \ 4) - ((Year \ 100) - (Year \ 400))
End Function

Private Function GetTotalMilliseconds(ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long) As Currency
    GetTotalMilliseconds = Hour * MILLISECONDS_PER_HOUR + Minute * MILLISECONDS_PER_MINUTE + Second * MILLISECONDS_PER_SECOND + Millisecond
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMilliseconds = .ReadProperty(PROP_MILLISECONDS, DEF_MILLISECONDS)
        mKind = .ReadProperty(PROP_KIND, DEF_KIND)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_MILLISECONDS, mMilliseconds)
        Call .WriteProperty(PROP_KIND, mKind)
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

Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByVal Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function

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
'    Module: HebrewCalendar
'

''
' Provides methods for manipulating Hebrew (Jewish) dates.
'
' @remarks Information about the Hebrew calendar was found at "http://astro.nmsu.edu/~lhuber/leaphist.html"
' and "http://www.geocities.com/Athens/1584/"
' @see HebrewCalendarStatic
' @see Calendar
'
Option Explicit
Implements IObject
Implements Calendar

Private Const PROP_TWODIGITYEARMAX As String = "TwoDigitYearMax"

' "http://listserv.brown.edu/archives/cgi-bin/wa?A2=ind0404d&L=conlang&F=&S=&P=19287"
' the Hebrew calendar epoch - that is, Tishri 1, 1 AM -
' corresponds to Monday, October 7, 3761 BC in the Julian calendar, which
' is September 7, 3761 BC in the Gregorian calendar, which is RD -1,373,427.
Private Const EPOCH As Long = -1373427

Private Const SECONDS_PER_MINUTE        As Long = 60
Private Const MINUTES_PER_HOUR          As Long = 60
Private Const HOURS_PER_DAY             As Long = 24

Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private Const PARTS_PER_HOUR            As Long = 1080

Private Const MIN_MILLISECONDS          As Currency = 49914748800000@
Private Const MAX_MILLISECONDS          As Currency = 70678396799999@
Private Const MIN_HEBREWYEAR            As Long = 5343
Private Const MAX_HEBREWYEAR            As Long = 6000

Private mTwoDigitYearMax As Long


Public Property Get Eras() As Long()
    Dim Ret(0) As Long
    Ret(0) = HebrewCalendar.HebrewEra
    Eras = Ret
End Property

Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

Public Property Let TwoDigitYearMax(ByVal RHS As Long)
    VerifyYear RHS
    mTwoDigitYearMax = RHS
End Property

Public Function AddDays(ByRef Time As Variant, ByVal Days As Long) As cDateTime
    Set AddDays = cDateTime.GetcDateTime(Time).AddDays(Days)
End Function

Public Function AddHours(ByRef Time As Variant, ByVal Hours As Long) As cDateTime
    Set AddHours = cDateTime.GetcDateTime(Time).AddHours(Hours)
End Function

Public Function AddMilliseconds(ByRef Time As Variant, ByVal Milliseconds As Double) As cDateTime
    Set AddMilliseconds = cDateTime.GetcDateTime(Time).AddMilliseconds(Milliseconds)
End Function

Public Function AddMinutes(ByRef Time As Variant, ByVal Minutes As Long) As cDateTime
    Set AddMinutes = cDateTime.GetcDateTime(Time).AddMinutes(Minutes)
End Function

Public Function AddMonths(ByRef Time As Variant, ByVal Months As Long) As cDateTime
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Dim TOD As Currency
    Dim ms As Currency
    
    If Months = 0 Then
        Set AddMonths = cDateTime.GetcDateTime(Time)
        Exit Function
    End If
    
    ms = cDateTime.GetcDateTime(Time).TotalMilliseconds
    TOD = Modulus(ms, MILLISECONDS_PER_DAY)
    Call GetDateParts(ms, Complete, Year, Month, Day)
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
    If Day > Max Then Day = Max
    
    Set AddMonths = ToDateTime(Year, Month, Day, 0, 0, 0, 0).AddMilliseconds(TOD)
End Function

Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As cDateTime
    Set AddSeconds = cDateTime.GetcDateTime(Time).AddSeconds(Seconds)
End Function

Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As cDateTime
    Set AddWeeks = AddDays(Time, Weeks * 7)
End Function

Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As cDateTime
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Dim ms As Currency
    
    ms = cDateTime.GetcDateTime(Time).TotalMilliseconds
    Call GetDateParts(ms, Complete, Year, Month, Day)
    Month = ToDotNetMonth(Year, Month)
    Year = Year + Years
    
    Set AddYears = ToDateTime(Year, Month, Day, 0, 0, 0, 0).AddMilliseconds(Modulus(ms, MILLISECONDS_PER_DAY))
End Function

Public Function GetEra(ByRef Time As Variant) As Long
    Call cDateTime.GetcDateTime(Time)   ' verifies we have a date.
    GetEra = HebrewCalendar.HebrewEra
End Function

Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = cDateTime.GetcDateTime(Time).Hour
End Function

Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = cDateTime.GetcDateTime(Time).Millisecond
End Function

Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = cDateTime.GetcDateTime(Time).Second
End Function

Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = cDateTime.GetcDateTime(Time).Minute
End Function

Public Function GetMonth(ByRef Time As Variant) As Long
    Call GetDateParts(cDateTime.GetcDateTime(Time).TotalMilliseconds, MonthPart, , GetMonth)
End Function

Public Function GetYear(ByRef Time As Variant) As Long
    Call GetDateParts(cDateTime.GetcDateTime(Time).TotalMilliseconds, YearPart, GetYear)
End Function

Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = InternalGetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
End Function

Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    If IsLeapYear(Year, Era) Then
        GetMonthsInYear = 13
    Else
        GetMonthsInYear = 12
    End If
End Function

Public Function GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyMonthYear(Month, Year)
    GetDaysInMonth = GetDaysInHebrewMonth(Year, ToHebrewMonth(Year, Month))
End Function

Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    GetDaysInYear = GetFixedDaysFromParts(Year + 1, 7, 1) - GetFixedDaysFromParts(Year, 7, 1)
End Function

Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    Call GetDateParts(cDateTime.GetcDateTime(Time).TotalMilliseconds, DayPart, , , GetDayOfMonth)
End Function

Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = cDateTime.GetcDateTime(Time).DayOfWeek
End Function

Public Function GetDayOfYear(ByRef Time As Variant) As Long
    Call GetDateParts(cDateTime.GetcDateTime(Time).TotalMilliseconds, DayOfTheYear, , , , GetDayOfYear)
End Function

Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    IsLeapYear = ((235 * Year - 234) Mod 19) > 11
End Function

Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyMonthYear(Month, Year)
    If Month = 7 Then IsLeapMonth = IsLeapYear(Year)
End Function

Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    If IsLeapYear(Year, Era) Then
        If Month = 6 Then IsLeapDay = (Day = 30)
    Else
        IsLeapDay = IsLeapMonth(Year, Month)
    End If
End Function

Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByRef Era As Variant) As cDateTime
    Call VerifyEra(Era)
    
    Dim d As Long
    d = GetDaysInMonth(Year, Month)
    
    If Day > d Then _
        Throw Cor.NewArgumentOutOfRangeException("Day", Message:=cString.Format("Day must be between 1 and {0} for month {1}.", d, Month))
    
    Dim Days As Long
    Days = GetFixedDaysFromParts(Year, ToHebrewMonth(Year, Month), Day) - 1
    Set ToDateTime = cDateTime.FromMilliseconds(Days * MILLISECONDS_PER_DAY + Hour * MILLISECONDS_PER_HOUR + Minute * MILLISECONDS_PER_MINUTE + Second * MILLISECONDS_PER_SECOND + Millisecond)
End Function

Public Function ToFourDigitYear(ByVal Year As Long) As Long
    Select Case Year
        Case Is < 0
            Throw Cor.NewArgumentOutOfRangeException("Year", Year, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
            
        Case Is < 100
            Dim Y As Long
            Y = Year Mod 100
            ToFourDigitYear = (mTwoDigitYearMax \ 100) * 100 + Y
            If Y > mTwoDigitYearMax Mod 100 Then ToFourDigitYear = ToFourDigitYear - 100
            
        Case Is < MIN_HEBREWYEAR, Is > MAX_HEBREWYEAR
            Throw Cor.NewArgumentOutOfRangeException("Year", Year, Environment.GetResourceString(ArgumentOutOfRange_Range, MIN_HEBREWYEAR, MAX_HEBREWYEAR))
            
        Case Else
            ToFourDigitYear = Year
    End Select
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
    If Year < MIN_HEBREWYEAR Or Year > MAX_HEBREWYEAR Then Throw Cor.NewArgumentOutOfRangeException("Year", Year, "Hebrew year must be between 5343 and 6000.")
End Sub

Private Sub VerifyMonthYear(ByVal Month As Long, ByVal Year As Long)
    Dim Max As Long
    
    If IsLeapYear(Year) Then
        Max = 13
    Else
        Max = 12
    End If
    
    If Month < 1 Or Month > Max Then _
        Throw Cor.NewArgumentOutOfRangeException("Month", Month, cString.Format("Month must be between 1 and {0} for year {1}.", Max, Year))
End Sub

' this is based on the algorithm found at "http://www.funaba.org/en/calendar.html"
Private Function GetDaysToHebrewYear(ByVal Year As Long) As Long
    Dim Months As Long
    Dim Parts As Long
    Dim Hours As Long
    Dim Days As Long
    Dim TotalDays As Long
    Dim r As Long
    
    ' months before the specified year.
    Months = ((Year * 235 - 234) \ 19)
    
    ' get the day and parts of a day in the remaining month.
    Days = Months \ PARTS_PER_HOUR
    r = Months - (Days * PARTS_PER_HOUR)
    Parts = (r * 793) + 204
    
    ' calculate the total days
    Hours = (Months * 12) + 11 + (Days * 793) + (Parts \ PARTS_PER_HOUR)
    TotalDays = (Months * 29) + (Hours \ 24)
    
    ' cannot start a year on a Sunday, Wednesday or Friday.
    If (TotalDays + 1) * 3 Mod 7 < 3 Then TotalDays = TotalDays + 1
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
        If Year1 - Year = 382 Then GetYearStartOffset = 1
    End If
End Function

''
' Calculations are performed using the month of 'Nissan' as the
' first month. This is the month that a Hebrew year actually
' increments by 1. The first month in Dot NET is Tishrei,
' so the calculated month needs to be adjusted to the Dot NET month.
'
' @param Year The year that contains the month to be mapped. The year
' may be a leap year which would alter the mapping.
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
' @param Year The year that contains the month to be mapped. The year
' may be a leap year which would alter the mapping.
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
        Case 2, 4, 6, 10, 13:   GetDaysInHebrewMonth = 29
        Case 8:                 If GetDaysInYear(Year) Mod 10 <> 5 Then GetDaysInHebrewMonth = 29
        Case 9:                 If GetDaysInYear(Year) Mod 10 = 3 Then GetDaysInHebrewMonth = 29
        Case 12:                If Not IsLeapYear(Year) Then GetDaysInHebrewMonth = 29
    End Select
End Function

Private Sub GetDateParts(ByVal ms As Currency, ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim FixedDays As Long
    
    If ms < MIN_MILLISECONDS Or ms > MAX_MILLISECONDS Then _
        Throw Cor.NewArgumentOutOfRangeException(ms, Message:="The Hebrew calendar only supports the years 5343 to 6000 (1582 to 2240 in Gregorian).")
    
    FixedDays = Int(ms / MILLISECONDS_PER_DAY) + 1
    
    Year = Int((FixedDays - EPOCH) / 365.246822205978)  ' this is an approximation
    Do While FixedDays >= GetFixedDaysFromParts(Year, 7, 1): Year = Year + 1: Loop
    Year = Year - 1
    
    If DatePart = YearPart Then Exit Sub
    
    If FixedDays < GetFixedDaysFromParts(Year, 1, 1) Then
        Month = 7
    Else
        Month = 1
    End If
    
    Do While FixedDays > GetFixedDaysFromParts(Year, Month, GetDaysInHebrewMonth(Year, Month)): Month = Month + 1: Loop
    If DatePart = MonthPart Then
        Month = ToDotNetMonth(Year, Month)
        Exit Sub
    End If
    
    Day = FixedDays - GetFixedDaysFromParts(Year, Month, 1) + 1
    If DatePart = DayPart Then Exit Sub
    
    DayOfYear = FixedDays - GetFixedDaysFromParts(Year, 7, 1) + 1
End Sub

Private Function GetFixedDays(ByRef Time As Variant) As Long
    Dim dt As cDateTime
    Set dt = cDateTime.GetcDateTime(Time)
    GetFixedDays = Int(dt.TotalMilliseconds / MILLISECONDS_PER_DAY) + 1
End Function

Private Function GetFixedDaysFromParts(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    Dim i As Long
    Dim Ret As Long
    Dim Max As Long
    
    Ret = EPOCH - 1 + GetDaysToHebrewYear(Year) + GetYearStartOffset(Year)
    
    If Month < 7 Then
        For i = 1 To Month - 1
            Ret = Ret + GetDaysInHebrewMonth(Year, i)
        Next i
        Max = GetMonthsInYear(Year)
    Else
        Max = Month - 1
    End If
    
    For i = 7 To Max
        Ret = Ret + GetDaysInHebrewMonth(Year, i)
    Next i
    
    GetFixedDaysFromParts = Ret + Day
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mTwoDigitYearMax = GetCalendarLong(CAL_HEBREW, CAL_ITWODIGITYEARMAX)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mTwoDigitYearMax = PropBag.ReadProperty(PROP_TWODIGITYEARMAX)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PROP_TWODIGITYEARMAX, mTwoDigitYearMax
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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


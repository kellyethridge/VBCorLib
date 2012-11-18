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
'    Module: HijriCalendar
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

Private Const DEF_TWODIGITYEARMAX       As Long = 1451

' Number of Gregorian days to 7/18/622.
' This is 2 days after the start of the Hijri calendar to
' allow for the Hijri Adjustment to still fit in the window.
Private Const EPOCH As Long = 227013

' The Hijri calendar has a 30 year cycle with 11 of those years being leap years.
' The common year is 354 days, leap years are 355.
' (19 * 354) + (11 * 355) = 10631
Private Const DAYS_PER_30YEARS          As Long = 10631

' This is used similar to adding 0.5 to a value and taking the integer to
' round to the nearest whole integer.
' DAYS_PER_30YEARS + (((30 * 355) - (30 * 354))/2) = 10646
Private Const YEAR_ROUNDING_BIAS        As Long = 10646

' This is 4/3/9666 in fixed days, the maximum valid Hijri date.
Private Const MAX_FIXED_DAYS            As Long = 3530161
Private Const MAX_YEAR                  As Long = 9666
Private Const MAX_SHORTMONTH            As Long = 4
Private Const MAX_SHORTDAY              As Long = 3
Private Const MONTHS_PER_YEAR           As Long = 12

Private Const SECONDS_PER_MINUTE        As Long = 60
Private Const MINUTES_PER_HOUR          As Long = 60
Private Const HOURS_PER_DAY             As Long = 24

Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private mHijriAdjustment    As Long
Private mTwoDigitYearMax    As Long
Private mMonthDays()        As Long



Public Property Get HijriAdjustment() As Long
    HijriAdjustment = mHijriAdjustment
End Property

Public Property Let HijriAdjustment(ByVal RHS As Long)
    If RHS < -2 Or RHS > 2 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, -2, 2), "HijriAdjustment", HijriAdjustment)
        
    mHijriAdjustment = RHS
End Property

Public Property Get Eras() As Long()
    Dim Ret(0) As Long
    Ret(0) = HijriCalendar.HijriEra
    Eras = Ret
End Property

Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

Public Property Let TwoDigitYearMax(ByVal RHS As Long)
    If RHS < 100 Or RHS > MAX_YEAR Then _
        Throw Cor.NewArgumentOutOfRangeException("Year must be between 100 and 9999 inclusively.", "TwoDigitYearMax", RHS)
    
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
    If Months < -120000 Or Months > 120000 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, -120000, 120000), "Months", Months)
    
    If Months = 0 Then
        Set AddMonths = cDateTime.GetcDateTime(Time)
        Exit Function
    End If
    
    Dim Year As Long
    Dim Month As Long
    Dim Day As Long
    Call GetDateParts(Time, Complete, Year, Month, Day)
    Month = Month + Months
    Year = Year + MathExt.DivRem(Month, MONTHS_PER_YEAR, Month)
    
    Dim MaxDays As Long
    MaxDays = GetDaysInMonth(Year, Month)
    If Day > MaxDays Then Day = MaxDays
    
    Dim FixedDays As Long
    FixedDays = GetFixedDaysFromParts(Year, Month, Day)
    
    Dim ms As Currency
    ms = cDateTime.GetcDateTime(Time).TotalMilliseconds
    Set AddMonths = cDateTime.FromMilliseconds(MILLISECONDS_PER_DAY * FixedDays + Modulus(ms, MILLISECONDS_PER_DAY))
End Function

Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As cDateTime
    Set AddSeconds = cDateTime.GetcDateTime(Time).AddSeconds(Seconds)
End Function

Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As cDateTime
    Set AddWeeks = cDateTime.GetcDateTime(Time).AddDays(Weeks * 7)
End Function

Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As cDateTime
    Set AddYears = AddMonths(Time, Years * MONTHS_PER_YEAR)
End Function

Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    Call GetDateParts(Time, DayPart, , , GetDayOfMonth)
End Function

Public Function GetDayOfYear(ByRef Time As Variant) As Long
    Call GetDateParts(Time, DayOfTheYear, , , , GetDayOfYear)
End Function

Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = cDateTime.GetcDateTime(Time).DayOfWeek
End Function

Public Function GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYearMonth(Year, Month)
    
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

Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    If IsLeapYear(Year, Era) Then
        GetDaysInYear = 355
    Else
        GetDaysInYear = 354
    End If
End Function

Public Function GetEra(ByRef Time As Variant) As Long
    cDateTime.GetcDateTime Time
    GetEra = 1
End Function

Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = cDateTime.GetcDateTime(Time).Hour
End Function

Public Function GetMilliseconds(ByRef Time As Variant) As Double
    GetMilliseconds = cDateTime.GetcDateTime(Time).Millisecond
End Function

Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = cDateTime.GetcDateTime(Time).Minute
End Function

Public Function GetMonth(ByRef Time As Variant) As Long
    Call GetDateParts(Time, MonthPart, , GetMonth)
End Function

Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    
    If Year = MAX_YEAR Then
        GetMonthsInYear = MAX_SHORTMONTH
    Else
        GetMonthsInYear = MONTHS_PER_YEAR
    End If
End Function

Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = cDateTime.GetcDateTime(Time).Second
End Function

Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = InternalGetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
End Function

Public Function GetYear(ByRef Time As Variant) As Long
    Call GetDateParts(Time, YearPart, GetYear)
End Function

Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    Dim MaxDays As Long
    
    MaxDays = Me.GetDaysInMonth(Year, Month)
    If Day < 1 Or Day > MaxDays Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, MaxDays), "Day", Day)
        
    If Day = 30 And Month = 12 Then IsLeapDay = IsLeapYear(Year, Era)
End Function

Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYearMonth(Year, Month)
    
    IsLeapMonth = False
End Function

Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    
    Select Case Year Mod 30
        Case 2, 5, 7, 10, 13, 16, 18, 21, 24, 26, 29: IsLeapYear = True
    End Select
End Function

Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByRef Era As Variant) As cDateTime
    Call VerifyEra(Era)
    Call VerifyYearMonthDay(Year, Month, Day)
    
    Dim ms As Currency
    ms = GetFixedDaysFromParts(Year, Month, Day) * MILLISECONDS_PER_DAY
    
    If ms >= 0 Then
        Set ToDateTime = cDateTime.FromMilliseconds(ms + Hour * MILLISECONDS_PER_HOUR + Minute * MILLISECONDS_PER_MINUTE + Second * MILLISECONDS_PER_SECOND + Millisecond)
    Else
        Throw Cor.NewArgumentOutOfRangeException("Invalid Year, Month, or Day for this calendar")
    End If
End Function

Public Function ToFourDigitYear(ByVal Year As Long) As Long
    Dim y As Long
    
    If Year < 0 Or Year > MAX_YEAR Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 9999), "Year", Year)

    If Year < 100 Then
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
    If Year < 1 Or Year > MAX_YEAR Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 9666), "Year", Year)
End Sub

Private Sub VerifyYearMonth(ByVal Year As Long, ByVal Month As Long)
    Call VerifyYear(Year)
    If (Month < 1) Or _
       (Month > MONTHS_PER_YEAR) Or _
       (Year = MAX_YEAR And Month > MAX_SHORTMONTH) Then
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, IIf(Year = 9666, 4, 12)), "Month", Month)
    End If
End Sub

Private Sub VerifyYearMonthDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long)
    Call VerifyYearMonth(Year, Month)
    If Year = MAX_YEAR Then
        If Month = MAX_SHORTMONTH Then
            If Day > MAX_SHORTDAY Then
                Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, MAX_SHORTDAY), "Day", Day)
            End If
        ElseIf Day > mMonthDays(Month - 1) Then
            Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, mMonthDays(Month - 1)), "Day", Day)
        End If
    End If
End Sub

Private Function GetFixedDays(ByRef Time As Variant) As Long
    Dim dt As cDateTime
    Set dt = cDateTime.GetcDateTime(Time)
    GetFixedDays = Int(dt.TotalMilliseconds / MILLISECONDS_PER_DAY) + mHijriAdjustment
    
    If GetFixedDays < 1 Or GetFixedDays > MAX_FIXED_DAYS Then _
        Throw Cor.NewArgumentOutOfRangeException("Hijri calendary only represents dates from 1/1/0001 to 4/3/9666.", "Time", Time)
End Function

Private Function GetFixedDaysFromParts(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    Dim Ret As Long
    Ret = (Year - 1) * 354              ' add days for each year, excluding first year.
    Ret = Ret + (Year * 11 + 3) \ 30    ' add additional days for leap years.
    Ret = Ret + (Month - 1) * 29        ' add the minimum of 29 days per month, excluding the first month.
    Ret = Ret + Month \ 2               ' add the 1 day for each of the 30 day months.
    
    GetFixedDaysFromParts = Ret + Day + EPOCH - 1
End Function

Private Sub GetDateParts(ByRef Time As Variant, ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim FixedDays As Long
    FixedDays = GetFixedDays(Time)
    
    Year = (((FixedDays - EPOCH) * 30) + YEAR_ROUNDING_BIAS) \ DAYS_PER_30YEARS
    If DatePart = YearPart Then Exit Sub
    
    DayOfYear = FixedDays - GetFixedDaysFromParts(Year, 1, 1) + 1
    If DatePart = DayOfTheYear Then Exit Sub
    
    If DayOfYear <> 355 Then
        Month = MathExt.Ceiling((DayOfYear - 30) / 29.5) + 1
    Else
        Month = 12
    End If
    
    If DatePart = MonthPart Then Exit Sub
    
    Day = FixedDays - GetFixedDaysFromParts(Year, Month, 1) + 1
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mTwoDigitYearMax = DEF_TWODIGITYEARMAX
    mMonthDays = Cor.NewLongs(0, 30, 59, 89, 118, 148, 177, 207, 236, 266, 295, 325, 355)
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mTwoDigitYearMax = .ReadProperty("TwoDigitYearMax")
        mHijriAdjustment = .ReadProperty("HijriAdjustment")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("TwoDigitYearMax", mTwoDigitYearMax)
        Call .WriteProperty("HijriAdjustment", mHijriAdjustment)
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



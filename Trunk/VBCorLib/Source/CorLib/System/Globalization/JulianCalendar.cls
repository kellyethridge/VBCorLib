VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JulianCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: JulianCalendar
'

''
' Provides methods to manipulate Julian dates.
'
' @see JulianCalendarStatic
' @see Calendar
'
Option Explicit
Implements IObject
Implements Calendar

Private Const DEF_TWODIGITYEARMAX       As Long = 2029
Private Const PROP_TWODIGITYEARMAX      As String = "TwoDigitYearMax"

Private Const SECONDS_PER_MINUTE        As Long = 60
Private Const MINUTES_PER_HOUR          As Long = 60
Private Const HOURS_PER_DAY             As Long = 24
Private Const MONTHS_PER_YEAR           As Long = 12

Private Const MILLISECONDS_PER_SECOND   As Currency = 1000@
Private Const MILLISECONDS_PER_MINUTE   As Currency = MILLISECONDS_PER_SECOND * SECONDS_PER_MINUTE
Private Const MILLISECONDS_PER_HOUR     As Currency = MILLISECONDS_PER_MINUTE * MINUTES_PER_HOUR
Private Const MILLISECONDS_PER_DAY      As Currency = MILLISECONDS_PER_HOUR * HOURS_PER_DAY

Private Const DAYS_PER_YEAR             As Long = 365
Private Const DAYS_PER_4_YEARS          As Long = DAYS_PER_YEAR * 4 + 1

Private mTwoDigitYearMax As Long



Public Property Get Eras() As Long()
    Dim Ret(0) As Long
    Ret(0) = 1
    Eras = Ret
End Property

Public Property Get TwoDigitYearMax() As Long
    TwoDigitYearMax = mTwoDigitYearMax
End Property

Public Property Let TwoDigitYearMax(ByVal RHS As Long)
    If RHS < 100 Or RHS > 9999 Then _
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
        Throw Cor.NewArgumentOutOfRangeException("Months must be between -120000 and 120000 inclusively.", "Months", Months)
    
    If Months = 0 Then
        Set AddMonths = cDateTime.GetcDateTime(Time)
        Exit Function
    End If

    Dim Month   As Long
    Dim Year    As Long
    Dim Day     As Long
    Call GetDateParts(Time, Complete, Year, Month, Day)
    
    Dim yearsToAdd As Long
    yearsToAdd = Int(Months / MONTHS_PER_YEAR)
    
    Dim monthsToAdd As Long
    monthsToAdd = Months - yearsToAdd * MONTHS_PER_YEAR
    
    Month = Month + monthsToAdd
    Year = Year + yearsToAdd
    If Month < 1 Then
        Month = Month + MONTHS_PER_YEAR
        Year = Year - 1
    ElseIf Month > MONTHS_PER_YEAR Then
        Month = Month - MONTHS_PER_YEAR
        Year = Year + 1
    End If
    
    Dim MaxDays As Long
    MaxDays = GetDaysInMonth(Year, Month)
    If Day > MaxDays Then Day = MaxDays
    
    Dim TOD As Currency
    TOD = cDateTime.GetcDateTime(Time).TotalMilliseconds
    
    Set AddMonths = cDateTime.FromMilliseconds(TotalDays(Year, Month, Day) * MILLISECONDS_PER_DAY + Modulus(TOD, MILLISECONDS_PER_DAY))
End Function

Public Function AddSeconds(ByRef Time As Variant, ByVal Seconds As Long) As cDateTime
    Set AddSeconds = cDateTime.GetcDateTime(Time).AddSeconds(Seconds)
End Function

Public Function AddWeeks(ByRef Time As Variant, ByVal Weeks As Long) As cDateTime
    Set AddWeeks = cDateTime.GetcDateTime(Time).AddDays(Weeks * 7)
End Function

Public Function AddYears(ByRef Time As Variant, ByVal Years As Long) As cDateTime
    Set AddYears = AddMonths(Time, Years * 12)
End Function

Public Function GetDayOfMonth(ByRef Time As Variant) As Long
    Call GetDateParts(Time, DayPart, , , GetDayOfMonth)
End Function

Public Function GetMonth(ByRef Time As Variant) As Long
    Call GetDateParts(Time, MonthPart, , GetMonth)
End Function

Public Function GetYear(ByRef Time As Variant) As Long
    Call GetDateParts(Time, YearPart, GetYear)
End Function

Public Function GetHour(ByRef Time As Variant) As Long
    GetHour = cDateTime.GetcDateTime(Time).Hour
End Function

Public Function GetMinute(ByRef Time As Variant) As Long
    GetMinute = cDateTime.GetcDateTime(Time).Minute
End Function

Public Function GetSecond(ByRef Time As Variant) As Long
    GetSecond = cDateTime.GetcDateTime(Time).Second
End Function

Public Function GetMilliseconds(ByRef Time As Variant) As Long
    GetMilliseconds = cDateTime.GetcDateTime(Time).Millisecond
End Function

Public Function GetDayOfWeek(ByRef Time As Variant) As DayOfWeek
    GetDayOfWeek = cDateTime.GetcDateTime(Time).DayOfWeek
End Function

Public Function GetDaysInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    If IsLeapYear(Year) Then
        GetDaysInYear = 366
    Else
        GetDaysInYear = 365
    End If
End Function

Public Function GetDayOfYear(ByRef Time As Variant) As Long
    Call GetDateParts(Time, DayOfTheYear, , , , GetDayOfYear)
End Function

Public Function GetEra(ByRef Time As Variant) As Long
    Call cDateTime.GetcDateTime(Time)
    GetEra = 1
End Function

Public Function GetMonthsInYear(ByVal Year As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    GetMonthsInYear = 12
End Function

Public Function GetDaysInMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Long
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
    GetDaysInMonth = cDateTime.DaysInMonth(Year, Month)
    If Month = 2 Then
        If IsLeapYear(Year) Then
            GetDaysInMonth = 29
        End If
    End If
End Function

Public Function IsLeapYear(ByVal Year As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    IsLeapYear = ((Year Mod 4) = 0)
End Function

Public Function IsLeapMonth(ByVal Year As Long, ByVal Month As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
    IsLeapMonth = False
End Function

Public Function IsLeapDay(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, Optional ByRef Era As Variant) As Boolean
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
    If Day < 1 Or Day > GetDaysInMonth(Year, Month, Era) Then _
        Throw Cor.NewArgumentOutOfRangeException(cString.Format("Day must be between 1 and {0} inclusively.", GetDaysInMonth(Year, Month, Era)), "Day", Day)
    
    If Month = 2 Then
        If IsLeapYear(Year, Era) Then
            IsLeapDay = (Day = 29)
        End If
    End If
End Function

Public Function ToDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long, Optional ByVal Era As Variant) As cDateTime
    Call VerifyEra(Era)
    Call VerifyYear(Year)
    Call VerifyMonth(Month)
        
    If Day < 1 Or Day > GetDaysInMonth(Year, Month) Then _
        Throw Cor.NewArgumentOutOfRangeException(cString.Format("Day must be between 1 and {0} inclusively.", GetDaysInMonth(Year, Month)), "Day", Day)
    
    Set ToDateTime = cDateTime.FromMilliseconds(TotalDays(Year, Month, Day) * MILLISECONDS_PER_DAY + TotalTime(Hour, Minute, Second, Millisecond))
End Function

Public Function ToFourDigitYear(ByVal Year As Long) As Long
    If Year < 0 Or Year > 9999 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 0, 9999), "Year", Year)

    If Year < 100 Then
        Dim y As Long
        y = Year Mod 100
        ToFourDigitYear = (mTwoDigitYearMax \ 100) * 100 + y
        If y > mTwoDigitYearMax Mod 100 Then ToFourDigitYear = ToFourDigitYear - 100
    Else
        ToFourDigitYear = Year
    End If
End Function

Public Function GetWeekOfYear(ByRef Time As Variant, ByVal Rule As CalendarWeekRule, ByVal FirstDayOfWeek As DayOfWeek) As Long
    GetWeekOfYear = InternalGetWeekOfYear(Time, Rule, FirstDayOfWeek, Me)
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
            If Era <> 1 Then Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidEraValue), "Era")
        Case Else
            Throw Cor.NewInvalidCastException("An integer value is required.")
    End Select
End Sub

Private Sub VerifyYear(ByVal Year As Long)
    If Year < 1 Or Year > 9999 Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 9999), "Year", Year)
End Sub

Private Sub VerifyMonth(ByVal Month As Long)
    If Month < 1 Or Month > 12 Then Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Range, 1, 12), "Month", Month)
End Sub

Private Sub GetDateParts(ByRef Time As Variant, ByVal DatePart As DatePartPrecision, Optional ByRef Year As Long, Optional ByRef Month As Long, Optional ByRef Day As Long, Optional ByRef DayOfYear As Long)
    Dim ms As Currency
    ms = cDateTime.GetcDateTime(Time).TotalMilliseconds
    
    Dim FixedDays As Long
    FixedDays = Int(ms / MILLISECONDS_PER_DAY) + 2  ' Julian calendar starts 2 days after Gregorian
    
    Dim years4 As Long
    years4 = FixedDays \ DAYS_PER_4_YEARS
    FixedDays = FixedDays - years4 * DAYS_PER_4_YEARS
    
    Dim Years As Long
    Years = FixedDays \ 365
    If Years = 4 Then Years = 3 ' leap year
    Year = years4 * 4 + Years + 1
    
    If DatePart = YearPart Then Exit Sub
    
    FixedDays = FixedDays - Years * DAYS_PER_YEAR
    DayOfYear = FixedDays + 1
    If DatePart = DayOfTheYear Then Exit Sub
    
    If Years = 3 Then
        Day = CalculateDays(FixedDays, Month, DaysToMonthLeapYear) + 1
    Else
        Day = CalculateDays(FixedDays, Month, DaysToMonth) + 1
    End If
End Sub

Private Function CalculateDays(ByVal Days As Long, ByRef m As Long, ByRef Months() As Long) As Long
    Do While Months(m) <= Days
        m = m + 1
    Loop
    CalculateDays = Days - Months(m - 1)
End Function

Private Function TotalDays(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Long
    If Year = 0 And Month = 0 And Day = 0 Then Exit Function
    
    If Year < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_ValidValues), "Year", Year)
    If Month < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_ValidValues), "Month", Month)
    If Day < 1 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_ValidValues), "Day", Day)
    
    Dim Days As Long
    If IsLeapYear(Year) Then
        Days = DaysToMonthLeapYear(Month - 1)
    Else
        Days = DaysToMonth(Month - 1)
    End If
    
    Year = Year - 1
    TotalDays = (Day - 1) + Days + 365 * Year + (Year \ 4) - 2 ' subtract 2 days to offset julian calendar start from gregorian
End Function

Private Function TotalTime(ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long, ByVal Millisecond As Long) As Currency
    TotalTime = Hour * MILLISECONDS_PER_HOUR + Minute * MILLISECONDS_PER_MINUTE + Second * MILLISECONDS_PER_SECOND + Millisecond
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mTwoDigitYearMax = DEF_TWODIGITYEARMAX
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    mTwoDigitYearMax = PropBag.ReadProperty(PROP_TWODIGITYEARMAX, DEF_TWODIGITYEARMAX)
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(PROP_TWODIGITYEARMAX, mTwoDigitYearMax)
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

Private Property Get Calendar_TwoDigitYearMax() As Long
    Calendar_TwoDigitYearMax = TwoDigitYearMax
End Property

Private Property Let Calendar_TwoDigitYearMax(ByVal RHS As Long)
    TwoDigitYearMax = RHS
End Property


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

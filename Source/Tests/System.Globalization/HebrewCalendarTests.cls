VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HebrewCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HebrewCalendarTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TwoDigitYearMax_WithValidMinMaxValues_SetsValue"
            Test.Use 5343
            Test.Use 5999
            Test.Use 99
            
        Case "ToFourDigitYear_WithTwoDigitYearMaxOf99_ReturnsExpected"
            Test.Use 1, 1
            Test.Use 29, 29
            Test.Use 30, 30
            Test.Use 99, 99
            Test.Use 5630, 5630
            
        Case "ToFourDigitYear_WithTwoDigitYearMaxOf5729_ReturnsExpected"
            Test.Use 30, 5630
            Test.Use 5630, 5630
            Test.Use 29, 5729
            
        Case "GetDaysInMonth_WithMonthInNonLeapYear_ReturnsExpected"
            Test.Use 1, 30
            Test.Use 2, 29
            Test.Use 3, 30
            Test.Use 4, 29
            Test.Use 5, 30
            Test.Use 6, 29
            Test.Use 7, 30
            Test.Use 8, 29
            Test.Use 9, 30
            Test.Use 10, 29
            Test.Use 11, 30
            Test.Use 12, 29
        
        Case "GetDaysInMonth_WithMonthInLeapYear_ReturnsExpected"
            Test.Use 1, 30
            Test.Use 2, 30
            Test.Use 3, 30
            Test.Use 4, 29
            Test.Use 5, 30
            Test.Use 6, 30
            Test.Use 7, 29
            Test.Use 8, 30
            Test.Use 9, 29
            Test.Use 10, 30
            Test.Use 11, 29
            Test.Use 12, 30
            Test.Use 13, 29
            
        Case "IsLeapYear_WithValues_ReturnsExpected"
            Test.Use 5555, False
            Test.Use 5763, True
            
        Case "IsLeapMonth_WithValues_ReturnsExpected"
            Test.Use 1, False
            Test.Use 2, False
            Test.Use 3, False
            Test.Use 4, False
            Test.Use 5, False
            Test.Use 6, False
            Test.Use 7, True
            Test.Use 8, False
            Test.Use 9, False
            Test.Use 10, False
            Test.Use 11, False
            Test.Use 12, False
            Test.Use 13, False
            
        Case "IsLeapDay_WithValues_ReturnsExpected"
            Test.Use 5763, 6, 29, False
            Test.Use 5763, 6, 30, True
            Test.Use 5763, 7, 1, True
            Test.Use 5763, 7, 29, True
            Test.Use 5763, 8, 1, False
            
        Case "GetWeekOfYear_WithValues_ReturnsExpected"
            Test.Use #1/1/2007#, CalendarWeekRule.FirstDay, DayOfWeek.Sunday, 16
            Test.Use #1/10/2000#, CalendarWeekRule.FirstDay, DayOfWeek.Sunday, 19
            Test.Use #1/10/2000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday, 18
            Test.Use #1/10/2000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Sunday, 18
            Test.Use #1/1/2000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday, 16
            Test.Use #1/1/2000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Sunday, 16
            
        Case "GetDaysInYear_WithYear_ReturnsExpected"
            Test.Use 5500, 355
            Test.Use 5502, 383
            Test.Use 5550, 353
            Test.Use 5343, 355
            Test.Use 5999, 385
            Test.Use 5763, 385
            Test.Use 5762, 354
            
        Case "AddYears_WithInvalidValue_ThrowsArgumentOutOfRangeException"
            Test.Use #12/31/1582 11:59:59 PM#
            Test.Use #9/30/2239#
            
    End Select
End Sub

Public Sub IsReadOnly_WhenNew_ReturnsFalse()
    Dim Calendar As New HebrewCalendar
    
    Assert.That Calendar.IsReadOnly, Iz.False
End Sub

Public Sub IsReadOnly_WhenMadeReadOnly_ReturnsTrue()
    Dim cal As New HebrewCalendar
    
    Dim Actual As Calendar
    Set Actual = CorLib.Calendar.ReadOnly(cal)
    
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub TestSerialization()
    Dim Calendar As New HebrewCalendar
    Calendar.TwoDigitYearMax = 99
    Set Calendar = HebrewCalendar.ReadOnly(Calendar)
    
    Dim Actual As Calendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.False
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(99)
End Sub

Public Sub TwoDigitYearMax_WithDefault_Returns5790()
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.TwoDigitYearMax
    
    Assert.That Actual, Iz.EqualTo(5790)
End Sub

Public Sub TwoDigitYearMax_WithMinYearMinusOne_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = 5342
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub TwoDigitYearMax_WithMaxYearPlusOne_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = 6000
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub TwoDigitYearMax_WithValidMinMaxValues_SetsValue(ByVal Year As Long)
    Dim Calendar As New HebrewCalendar
    
    Calendar.TwoDigitYearMax = Year
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(Year)
End Sub

Public Sub ToFourDigitYear_WithTwoDigitYearMaxOf99_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    Calendar.TwoDigitYearMax = 99
    
    Dim Actual  As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToFourDigitYear_WithTwoDigitYearMaxOf5729_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    Calendar.TwoDigitYearMax = 5729
    
    Dim Actual  As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToFourDigitYear_WithLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.ToFourDigitYear -1
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub Eras_WhenCalled_ReturnsArrayOfOne()
    Dim Calendar As New HebrewCalendar
    
    Dim Actual() As Long
    Actual = Calendar.Eras
    
    Assert.That Actual, Iz.EqualTo(NewLongs(1))
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsDateTimeMinValue()
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MinSupportedDateTime
    
    Assert.That Actual, Equals(#1/1/1583#)
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsDateTimeMaxValue()
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MaxSupportedDateTime
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2239, 9, 29, 23, 59, 59, 999)
    Set Expected = CorDateTime.FromMilliseconds(Expected.TotalMilliseconds + 0.9999@)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub GetProperties_WhenCalled_ReturnsExpected()
    Const d As Date = #1/2/2200 1:20:35 PM#
    Dim Calendar As New HebrewCalendar
    
    Assert.That Calendar.GetYear(d), Iz.EqualTo(5960), "#1"
    Assert.That Calendar.GetMonth(d), Iz.EqualTo(4), "#2"
    Assert.That Calendar.GetDayOfMonth(d), Iz.EqualTo(14), "#3"
    Assert.That Calendar.GetHour(d), Iz.EqualTo(13), "#4"
    Assert.That Calendar.GetMinute(d), Iz.EqualTo(20), "#5"
    Assert.That Calendar.GetSecond(d), Iz.EqualTo(35), "#6"
    Assert.That Calendar.GetDayOfWeek(d), Iz.EqualTo(DayOfWeek.Thursday), "#7"
    Assert.That Calendar.GetDayOfYear(d), Iz.EqualTo(104), "#8"
End Sub

Public Sub GetMilliseconds_WhenCalled_ReturnsMillisecondsPart()
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetMilliseconds(NewDateTime(Second:=1, Millisecond:=50))
    
    Assert.That Actual, Iz.EqualTo(50)
End Sub

Public Sub AddMethods_WhenCalled_ReturnsExpected()
    Const d As Date = #1/1/1800#
    Dim Calendar As New HebrewCalendar

    Assert.That Calendar.AddYears(d, 1), Equals(#12/20/1800#), "#1"
    Assert.That Calendar.AddMonths(d, 1), Equals(#1/30/1800#), "#2"
    Assert.That Calendar.AddMonths(d, -1), Equals(#12/2/1799#), "#3"
    Assert.That Calendar.AddDays(d, 1), Equals(#1/2/1800#), "#4"
    Assert.That Calendar.AddHours(d, 1), Equals(#1/1/1800 1:00:00 AM#), "#5"
    Assert.That Calendar.AddMinutes(d, 1), Equals(#1/1/1800 12:01:00 AM#), "#6"
    Assert.That Calendar.AddSeconds(d, 1), Equals(#1/1/1800 12:00:01 AM#), "#7"
    Assert.That Calendar.AddMilliseconds(d, 1), Equals(NewDateTime(1800, 1, 1, Millisecond:=1)), "#8"
    Assert.That Calendar.AddWeeks(d, 5), Equals(#2/5/1800#), "#9"
End Sub

Public Sub AddYears_WithInvalidValue_ThrowsArgumentOutOfRangeException(ByVal Time As Date)
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.AddYears Time, 1
    
    AssertArgumentOutOfRangeException Err, "Time"
End Sub

Public Sub Kind_WhenAddMethodCalled_ReturnsUnspecified()
    Dim Calendar As New HebrewCalendar
    Dim d As CorDateTime
    Set d = NewDate(#1/1/1800#, UtcKind)
    
    Assert.That Calendar.AddDays(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#1"
    Assert.That Calendar.AddMonths(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#2"
    Assert.That Calendar.AddYears(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#3"
    Assert.That Calendar.AddHours(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#4"
    Assert.That Calendar.AddMinutes(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#5"
    Assert.That Calendar.AddSeconds(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#6"
    Assert.That Calendar.AddMilliseconds(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#7"
    Assert.That Calendar.AddWeeks(d, 1).Kind, Iz.EqualTo(DateTimeKind.UnspecifiedKind), "#8"
End Sub

Public Sub GetDaysInMonth_WithMonthInNonLeapYear_ReturnsExpected(ByVal Month As Long, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(5555, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInYear_WithYear_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_WithMonthInLeapYear_ReturnsExpected(ByVal Month As Long, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(5763, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_MonoBug76252()
    Dim Calendar    As New HebrewCalendar
    Dim Year        As Long
    
    Year = Calendar.GetYear(#9/1/2005#)
    Assert.That Year, Iz.EqualTo(5765)
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(Year, 13, 1)
    
    Assert.That Actual, Iz.EqualTo(29)
End Sub

Public Sub IsLeapYear_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapYear_With99_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.IsLeapYear 99
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub IsLeapMonth_WithValues_ReturnsExpected(ByVal Month As Long, ByVal Expected As Boolean)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapMonth(5763, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapDay_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As Boolean)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapDay(Year, Month, Day)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapDay_WithInvalidDayValue_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New HebrewCalendar
    
    On Error Resume Next
    Calendar.IsLeapDay 5343, 1, 50
    
    AssertArgumentOutOfRangeException Err, "Day"
End Sub

Public Sub GetWeekOfYear_WithValues_ReturnsExpected(ByVal d As Date, ByVal Rule As CalendarWeekRule, ByVal dow As DayOfWeek, ByVal Expected As Long)
    Dim Calendar As New HebrewCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetWeekOfYear(d, Rule, dow)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

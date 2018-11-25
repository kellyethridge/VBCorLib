VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GregorianCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' GregorianCalendarTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToFourDigitYear_WithTwoDigitYearMaxOf99_ReturnsExpected"
            Test.Use 1, 1
            Test.Use 29, 29
            Test.Use 30, 30
            Test.Use 99, 99
            Test.Use 100, 100
        
        Case "ToFourDigitYear_WithTwoDigitYearMaxOf101_ReturnsExpected"
            Test.Use 1, 101
            Test.Use 2, 2
            Test.Use 101, 101
        
        Case "GetDaysInMonth_WithMonthValue_ReturnsExpected"
            Test.Use 1, 31
            Test.Use 2, 28
            Test.Use 3, 31
            Test.Use 4, 30
            Test.Use 5, 31
            Test.Use 6, 30
            Test.Use 7, 31
            Test.Use 8, 31
            Test.Use 9, 30
            Test.Use 10, 31
            Test.Use 11, 30
            Test.Use 12, 31
            
        Case "IsLeapYear_WithValues_ReturnsExpected"
            Test.Use 2017, False
            Test.Use 2016, True
            Test.Use 2015, False
            
        Case "IsLeapMonth_WithValues_ReturnsExpected"
            Test.Use 2017, 1
            Test.Use 2017, 2
            Test.Use 2016, 1
            Test.Use 2016, 2
        
        Case "IsLeapDay_WithValues_ReturnsExpected"
            Test.Use 2017, 2, 28, False
            Test.Use 2016, 2, 29, True
            
        Case "GetWeekOfYear_WithValues_ReturnsExpected"
            Test.Use #1/1/2007#, CalendarWeekRule.FirstDay, DayOfWeek.Sunday, 1
            Test.Use #1/10/2000#, CalendarWeekRule.FirstDay, DayOfWeek.Sunday, 3
            Test.Use #1/10/2000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday, 2
            Test.Use #1/10/2000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Sunday, 2
            Test.Use #1/1/2000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday, 52
            Test.Use #1/1/2000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Sunday, 52
        
    End Select
End Sub

Public Sub AlgorithmType_WhenCalled_ReturnsSolarCalendar()
    Dim Calendar As New GregorianCalendar
    
    Assert.That Calendar.AlgorithmType, Iz.EqualTo(CalendarAlgorithmType.SolarCalendar)
End Sub

Public Sub Clone_WhenCalled_ReturnsWritableClone()
    Dim Calendar As GregorianCalendar
    Set Calendar = Cor.NewGregorianCalendar(MiddleEastFrench)
    Calendar.TwoDigitYearMax = 5400
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Dim Actual As GregorianCalendar
    Set Actual = Calendar.Clone
    
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(5400)
    Assert.That Actual.CalendarType, Iz.EqualTo(GregorianCalendarTypes.MiddleEastFrench)
    Assert.That Actual.IsReadOnly, Iz.False
End Sub

Public Sub IsReadOnly_WhenNew_ReturnsFalse()
    Dim Calendar As New GregorianCalendar
    
    Assert.That Calendar.IsReadOnly, Iz.False
End Sub

Public Sub IsReadOnly_WhenMadeReadOnly_ReturnsTrue()
    Dim cal As New GregorianCalendar
    
    Dim Actual As Calendar
    Set Actual = CorLib.Calendar.ReadOnly(cal)
    
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub TwoDigitYearMax_WithDefault_Returns2029()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.TwoDigitYearMax
    
    Assert.That Actual, Iz.EqualTo(2029)
End Sub

Public Sub TwoDigitYearMax_WithMinYearMinusOne_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New GregorianCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = 98
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub TwoDigitYearMax_WithMaxYearPlusOne_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New GregorianCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = 10000
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub TwoDigitYearMax_WhenReadOnly_ThrowsInvalidOperationException()
    Dim Calendar As New GregorianCalendar
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = 99
    
    AssertInvalidOperationException Err
End Sub

Public Sub CalendarType_WhenReadOnly_ThrowsInvalidOperationException()
    Dim Calendar As New GregorianCalendar
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    On Error Resume Next
    Calendar.CalendarType = Arabic
    
    AssertInvalidOperationException Err
End Sub

Public Sub ToFourDigitYear_WithTwoDigitYearMaxOf99_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New GregorianCalendar
    Calendar.TwoDigitYearMax = 99
    
    Dim Actual  As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToFourDigitYear_WithTwoDigitYearMaxOf101_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New GregorianCalendar
    Calendar.TwoDigitYearMax = 101
    
    Dim Actual As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Eras_WhenCalled_ReturnsArrayOfOne()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual() As Long
    Actual = Calendar.Eras
    
    Assert.That Actual, Iz.EqualTo(NewLongs(1))
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsDateTimeMinValue()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MinSupportedDateTime
    
    Assert.That Actual, Equals(CorDateTime.MinValue)
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsDateTimeMaxValue()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MaxSupportedDateTime
    
    Assert.That Actual, Equals(CorDateTime.MaxValue)
End Sub

Public Sub TestSerialization()
    Dim Calendar As New GregorianCalendar
    Calendar.TwoDigitYearMax = 99
    Calendar.CalendarType = Arabic
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Dim Actual As GregorianCalendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.False
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(99)
    Assert.That Actual.CalendarType, Iz.EqualTo(GregorianCalendarTypes.Arabic)
End Sub

Public Sub GetProperties_WhenCalled_ReturnsExpected()
    Const d As Date = #1/2/2001 1:20:35 PM#
    Dim Calendar As New GregorianCalendar
    
    Assert.That Calendar.GetYear(d), Iz.EqualTo(2001), "#1"
    Assert.That Calendar.GetMonth(d), Iz.EqualTo(1), "#2"
    Assert.That Calendar.GetDayOfMonth(d), Iz.EqualTo(2), "#3"
    Assert.That Calendar.GetHour(d), Iz.EqualTo(13), "#4"
    Assert.That Calendar.GetMinute(d), Iz.EqualTo(20), "#5"
    Assert.That Calendar.GetSecond(d), Iz.EqualTo(35), "#6"
    Assert.That Calendar.GetDayOfWeek(d), Iz.EqualTo(DayOfWeek.Tuesday), "#7"
End Sub

Public Sub GetMilliseconds_WhenCalled_ReturnsMillisecondsPart()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetMilliseconds(NewDateTime(Second:=1, Millisecond:=50))
    
    Assert.That Actual, Iz.EqualTo(50)
End Sub

Public Sub AddMethods_WhenCalled_ReturnsExpected()
    Const d As Date = #1/1/100#
    Dim Calendar As New GregorianCalendar

    Assert.That Calendar.AddYears(d, 1), Equals(NewDate(#1/1/101#)), "#1"
    Assert.That Calendar.AddMonths(d, 1), Equals(NewDate(#2/1/100#)), "#2"
    Assert.That Calendar.AddDays(d, 1), Equals(NewDate(#1/2/100#)), "#3"
    Assert.That Calendar.AddHours(d, 1), Equals(NewDate(#1/1/100 1:00:00 AM#)), "#4"
    Assert.That Calendar.AddMinutes(d, 1), Equals(NewDate(#1/1/100 12:01:00 AM#)), "#5"
    Assert.That Calendar.AddSeconds(d, 1), Equals(NewDate(#1/1/100 12:00:01 AM#)), "#6"
    Assert.That Calendar.AddMilliseconds(d, 1), Equals(NewDateTime(100, 1, 1, Millisecond:=1)), "#7"
    Assert.That Calendar.AddWeeks(d, 5), Equals(NewDate(#2/5/100#)), "#8"
End Sub

Public Sub AddYears_WithInvalidResult_ThrowsArgumentException()
    Dim Calendar As New GregorianCalendar
    
    On Error Resume Next
    Calendar.AddYears #1/1/2000#, 9000
    
    AssertArgumentException Err
End Sub

Public Sub Kind_WhenAddMethodCalled_ReturnsUnspecified()
    Dim Calendar As New GregorianCalendar
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

Public Sub GetDaysInMonth_WithMonthValue_ReturnsExpected(ByVal Month As Long, ByVal Expected As Long)
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(2017, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_WithFebuaryInLeapYear_Returns29()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(2016, 2)
    
    Assert.That Actual, Iz.EqualTo(29)
End Sub

Public Sub GetDaysInYear_WhenNonLeapYear_Returns365()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInYear(2018)
    
    Assert.That Actual, Iz.EqualTo(365)
End Sub

Public Sub GetDaysInYear_WhenIsLeapYear_Returns366()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInYear(2004)
    
    Assert.That Actual, Iz.EqualTo(366)
End Sub

Public Sub IsLeapYear_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapMonth_WhenIsLeapYear_ReturnsFalse()
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapMonth(2004, 2)
    
    Assert.That Actual, Iz.False
End Sub

Public Sub IsLeapDay_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As Boolean)
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapDay(Year, Month, Day)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetWeekOfYear_WithValues_ReturnsExpected(ByVal d As Date, ByVal Rule As CalendarWeekRule, ByVal dow As DayOfWeek, ByVal Expected As Long)
    Dim Calendar As New GregorianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetWeekOfYear(d, Rule, dow)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub














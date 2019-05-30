VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaiwanCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TaiwanCalendarTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TwoDigitYearMax_WithBoundryValue_SetsValue"
            Test.Use 99
            Test.Use 8088
    
        Case "TwoDigitYearMax_WhenSetWithInvalidValue_ThrowsArgumentOutOfRangeException"
            Test.Use 98
            Test.Use 8089
            
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
            Test.Use 92, False
            Test.Use 93, True
            Test.Use 94, False
            
        Case "IsLeapDay_WithValues_ReturnsExpected"
            Test.Use 92, 2, 28, False
            Test.Use 93, 2, 29, True
            
        Case "GetDaysInYear_WithYear_ReturnsExpected"
            Test.Use 92, 365
            Test.Use 93, 366
            
        Case "ToDateTime_WhenCalledWithDateTimeBoundry_ReturnsGregorianDateTime"
            Test.Use 1, 1, 1, NewDateTime(1912, 1, 1)
            Test.Use 8088, 12, 31, NewDateTime(9999, 12, 31)
            
        Case "ToDateTime_WithValueOutOfSupportedRange_ThrowsArgumentOutOfRangeException"
            Test.Use 0, 12, 31
            Test.Use 8089, 1, 1
        
        Case "ToFourDigitYear_WithValue_ReturnsExpected"
            Test.Use 1, 1
            Test.Use 8088, 8088
            
    End Select
End Sub

Public Sub AlgorithmType_WhenCalled_ReturnsSolarCalendar()
    Dim Calendar As New TaiwanCalendar
    
    Assert.That Calendar.AlgorithmType, Iz.EqualTo(CalendarAlgorithmType.SolarCalendar)
End Sub

Public Sub IsReadOnly_WhenCreated_ReturnsFalse()
    Dim Calendar As New TaiwanCalendar
    
    Assert.That Calendar.IsReadOnly, Iz.False
End Sub

Public Sub Era_WhenCalled_ReturnsOneEra()
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual() As Long
    Actual = Calendar.Eras
    
    Assert.That Actual, Iz.EqualTo(NewLongs(1))
End Sub

Public Sub TwoDigitYearMax_WhenNew_Returns99()
    Dim Calendar As New TaiwanCalendar
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(99)
End Sub

Public Sub TwoDigitYearMax_WithBoundryValue_SetsValue(ByVal Value As Long)
    Dim Calendar As New TaiwanCalendar
    
    Calendar.TwoDigitYearMax = Value
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(Value)
End Sub

Public Sub TwoDigitYearMax_WhenSetWithInvalidValue_ThrowsArgumentOutOfRangeException(ByVal InvalidValue As Long)
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = InvalidValue
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsDateTimeMinValue()
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MinSupportedDateTime
    
    Assert.That Actual, Equals(#1/1/1912#)
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsDateTimeMaxValue()
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MaxSupportedDateTime
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(9999, 12, 31, 23, 59, 59, 999).AddTicks(9999)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMethods_WhenCalled_ReturnsExpected()
    Const d As Date = #1/1/2900#
    Dim Calendar As New TaiwanCalendar

    Assert.That Calendar.AddYears(d, 1), Equals(#1/1/2901#), "AddYears #1"
    Assert.That Calendar.AddYears(d, -1), Equals(#1/1/2899#), "AddYears #2"
    Assert.That Calendar.AddMonths(d, 1), Equals(#2/1/2900#), "AddMonths #1"
    Assert.That Calendar.AddMonths(d, -1), Equals(#12/1/2899#), "AddMonths #2"
    Assert.That Calendar.AddDays(d, 1), Equals(#1/2/2900#), "AddDays #1"
    Assert.That Calendar.AddDays(d, -1), Equals(#12/31/2899#)
    Assert.That Calendar.AddHours(d, 1), Equals(#1/1/2900 1:00:00 AM#), "AddHours #1"
    Assert.That Calendar.AddMinutes(d, 1), Equals(#1/1/2900 12:01:00 AM#), "AddMinutes #1"
    Assert.That Calendar.AddSeconds(d, 1), Equals(#1/1/2900 12:00:01 AM#), "AddSeconds #1"
    Assert.That Calendar.AddMilliseconds(d, 1), Equals(NewDateTime(2900, 1, 1, Millisecond:=1)), "AddMilliseconds #1"
    Assert.That Calendar.AddWeeks(d, 5), Equals(#2/5/2900#), "AddWeeks #1"
    Assert.That Calendar.AddWeeks(d, -5), Equals(#11/27/2899#), "AddWeeks #2"
End Sub

Public Sub AddMilliseconds_WhenResultIsLessThanMinSupported_ThrowsArgumentException()
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.AddMilliseconds Calendar.MinSupportedDateTime, -1
    
    AssertArgumentException Err
End Sub

Public Sub AddMilliseconds_WhenResultGreaterThanMaxSupported_ThrowsArgumentException()
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.AddMilliseconds Calendar.MaxSupportedDateTime, 1
    
    AssertArgumentException Err
End Sub

Public Sub AddMonths_WhenResultIsLessThanMinSupported_ThrowsArgumentException()
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.AddMonths Calendar.MinSupportedDateTime, -1
    
    AssertArgumentException Err
End Sub

Public Sub AddMonths_WhenResultGreaterThanMaxSupported_ThrowsArgumentException()
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.AddMonths Calendar.MaxSupportedDateTime, 1
    
    AssertArgumentException Err
End Sub

Public Sub GetProperties_WhenCalled_ReturnsExpected()
    Dim d As CorDateTime
    Dim Calendar As New TaiwanCalendar
    
    Set d = NewDateTime(2200, 1, 2, 13, 20, 35, 100)
    
    Assert.That Calendar.GetYear(d), Iz.EqualTo(289), "#1"
    Assert.That Calendar.GetMonth(d), Iz.EqualTo(1), "#2"
    Assert.That Calendar.GetDayOfMonth(d), Iz.EqualTo(2), "#3"
    Assert.That Calendar.GetHour(d), Iz.EqualTo(13), "#4"
    Assert.That Calendar.GetMinute(d), Iz.EqualTo(20), "#5"
    Assert.That Calendar.GetSecond(d), Iz.EqualTo(35), "#6"
    Assert.That Calendar.GetMilliseconds(d), Iz.EqualTo(100), "#7"
    Assert.That Calendar.GetDayOfWeek(d), Iz.EqualTo(DayOfWeek.Thursday), "#8"
    Assert.That Calendar.GetDayOfYear(d), Iz.EqualTo(2), "#9"
End Sub

Public Sub GetDaysInMonth_WithMonthValue_ReturnsExpected(ByVal Month As Long, ByVal Expected As Long)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(1912, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_WithFebuaryInLeapYear_Returns29()
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(93, 2)
    
    Assert.That Actual, Iz.EqualTo(29)
End Sub

Public Sub GetDaysInYear_WithYear_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapYear_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapDay_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As Boolean)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapDay(Year, Month, Day)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToDateTime_WhenCalledWithDateTimeBoundry_ReturnsGregorianDateTime(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As CorDateTime)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.ToDateTime(Year, Month, Day, 0, 0, 0, 0)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub ToDateTime_WithValueOutOfSupportedRange_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long)
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.ToDateTime Year, Month, Day, 0, 0, 0, 0
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToFourDigitYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New TaiwanCalendar
    
    Dim Actual As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToFourDigitYear_WithInvalidValue_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New TaiwanCalendar
    
    On Error Resume Next
    Calendar.ToFourDigitYear 8089
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub Clone_WhenCalled_ReturnsWritableClone()
    Dim Calendar As New TaiwanCalendar
    Calendar.TwoDigitYearMax = 5400
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)

    Dim Actual As Calendar
    Set Actual = Calendar.Clone

    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(5400)
    Assert.That Actual.IsReadOnly, Iz.False
End Sub

Public Sub IsReadOnly_WhenMadeReadOnly_ReturnsTrue()
    Dim cal As New TaiwanCalendar
    
    Dim Actual As Calendar
    Set Actual = CorLib.Calendar.ReadOnly(cal)
    
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub TestSerialization()
    Dim Calendar As New TaiwanCalendar
    Calendar.TwoDigitYearMax = 5400
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Dim Actual As Calendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.False
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(5400)
End Sub

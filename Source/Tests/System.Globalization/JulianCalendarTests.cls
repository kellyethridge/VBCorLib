VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JulianCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' JulianCalendarTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "TwoDigitYearMax_WithInvalidValue_ThrowsArgumentOutOfRangeException"
            Test.Use 98
            Test.Use 10000
        
        Case "AddMonths_WithValue_ReturnsExpected"
            Test.Use #1/6/1000#, -1, #12/6/999#
            Test.Use #12/31/9999#, -119985, NewDateTime(1, 1, 17)
            Test.Use #1/31/1000#, 1, #3/3/1000#
            
        Case "GetYear_WithValue_ReturnsExpected"
            Test.Use #3/3/1000#, 1000
            Test.Use #1/1/1000#, 999
            
        Case "GetDayOfYear_WithValue_ReturnsExpected"
            Test.Use #1/1/1000#, 361
            Test.Use #1/5/1000#, 365
            Test.Use #1/6/1000#, 1
            Test.Use #12/31/1001#, 359
            
        Case "GetWeekOfYear_WithValues_ReturnsExpected"
            Test.Use #1/1/1000#, CalendarWeekRule.FirstDay, DayOfWeek.Monday, 53
            Test.Use #1/1/1000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday, 52
            Test.Use #1/1/1000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Monday, 52
            Test.Use #1/1/1000#, CalendarWeekRule.FirstDay, DayOfWeek.Friday, 52
            Test.Use #1/1/1000#, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Friday, 52
            Test.Use #1/1/1000#, CalendarWeekRule.FirstFullWeek, DayOfWeek.Friday, 51
            
        Case "IsLeapYear_WithValue_ReturnsExpected"
            Test.Use 1600, True
            Test.Use 1599, False
            Test.Use 1601, False
            
        Case "ToDateTime_WithInvalidDay_ThrowsArgumentOutOfRangeException"
            Test.Use 1600, 2, 0
            Test.Use 1600, 2, 30
            Test.Use 1600, 12, 0
            Test.Use 1600, 12, 32
            Test.Use 1, 1, 2
            
        Case "ToFourDigitYear_WithValue_ReturnsExpected"
            Test.Use 0, 2000
            Test.Use 1, 2001
            Test.Use 99, 1999
            Test.Use 100, 100
            Test.Use 9999, 9999
            
        Case "GetDaysInMonth_WithValue_ReturnsExpected"
            Test.Use 1599, 2, 28
            Test.Use 1600, 2, 29
            Test.Use 1599, 1, 31
            Test.Use 1600, 1, 31
            Test.Use 2004, 2, 29
            Test.Use 1000, 2, 28
            
        Case "IsLeapDay_WithValue_ReturnsExpected"
            Test.Use #2/28/1600#, False
            Test.Use #2/29/1600#, True
            Test.Use #3/29/1600#, False
            
    End Select
End Sub

Public Sub JulianEra_WhenCalled_ReturnsOne()
    Dim Actual As Long
    
    Actual = JulianCalendar.JulianEra
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub Eras_WhenCalled_ReturnsSingleElementArrayWithValueOne()
    Dim Calendar As New JulianCalendar
    
    Dim Actual() As Long
    Actual = Calendar.Eras
    
    Assert.That Actual, Iz.EqualTo(NewLongs(1))
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsExpected()
    Dim Calendar As New JulianCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MaxSupportedDateTime
    
    Assert.That Actual, Equals(CorDateTime.MaxValue)
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsExpected()
    Dim Calendar As New JulianCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MinSupportedDateTime
    
    Assert.That Actual, Equals(CorDateTime.MinValue)
End Sub

Public Sub TwoDigitYearMax_WithInvalidValue_ThrowsArgumentOutOfRangeException(ByVal InvalidValue As Long)
    Dim Calendar As New JulianCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = InvalidValue
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub TwoDigitYearMax_WithMinValidValue_SetsValue()
    Dim Calendar As New JulianCalendar
    
    Calendar.TwoDigitYearMax = 99
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(99)
End Sub

Public Sub TwoDigitYearMax_WithMaxValidValue_SetsValue()
    Dim Calendar As New JulianCalendar
    
    Calendar.TwoDigitYearMax = 9999
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(9999)
End Sub

Public Sub TestSerialization()
    Dim Calendar As New JulianCalendar
    Calendar.TwoDigitYearMax = 500
    
    Dim Actual As Calendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.False
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(500)
End Sub

Public Sub AddMonths_WithMultipleCalls_ReturnsExpected()
    Dim Calendar    As New JulianCalendar
    Dim Actual      As CorDateTime
    Dim i           As Long

    Set Actual = CorDateTime.FromOADate(#1/31/1000#)
    
    For i = -100 To 100
        Set Actual = Calendar.AddMonths(Actual, i)
    Next i
    
    Assert.That Actual, Equals(#1/31/1000#)
End Sub

Public Sub AddMonths_WithValue_ReturnsExpected(ByVal StartDate As Date, ByVal MonthsToAdd As Long, ByVal Expected As Variant)
    Dim Calendar    As New JulianCalendar
    Dim Actual      As CorDateTime
    
    Set Actual = Calendar.AddMonths(StartDate, MonthsToAdd)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub GetDayOfMonth_WhenCalled_ReturnsExpected()
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDayOfMonth(#3/3/1000#)
    
    Assert.That Actual, Iz.EqualTo(26)
End Sub

Public Sub GetDaysInMonth_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Expected As Long)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(Year, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetMonth_WhenCalled_ReturnsExpected()
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetMonth(#3/3/1000#)
    
    Assert.That Actual, Iz.EqualTo(2)
End Sub

Public Sub GetYear_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetYear(d)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDayOfYear_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDayOfYear(d)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToDateTime_WhenCalled_ReturnGregorianDateTime()
    Dim Calendar As New JulianCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.ToDateTime(1000, 2, 28, 1, 1, 1, 0)
    
    Assert.That Actual, Equals(#3/5/1000 1:01:01 AM#)
End Sub

Public Sub GetWeekOfYear_WithValues_ReturnsExpected(ByVal d As Date, ByVal Rule As CalendarWeekRule, ByVal FirstDay As DayOfWeek, ByVal Expected As Long)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetWeekOfYear(d, Rule, FirstDay)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapDay_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Boolean)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Boolean
    Actual = Calendar.IsLeapDay(Year(d), Month(d), Day(d))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToDateTime_WithInvalidDay_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long)
    Dim Calendar As New JulianCalendar
    
    On Error Resume Next
    Calendar.ToDateTime Year, Month, Day, 1, 1, 1, 1
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToFourDigitYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New JulianCalendar
    
    Dim Actual As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ReadOnly_WhenCalled_MarksCalendarReadOnly()
    Dim Calendar As New JulianCalendar
    
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Assert.That Calendar.IsReadOnly, Iz.True
End Sub


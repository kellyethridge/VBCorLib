VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JapaneseCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' JapaneseCalendarTests
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
            Const MinMinusOne As Long = 98
            Const MaxPlusOne As Long = 8012
            Test.Use MinMinusOne
            Test.Use MaxPlusOne
        
        Case "GetYear_WithEraBoundryDate_ReturnsExpected"
            Test.Use #9/8/1868#, 1
            Test.Use #7/29/1912#, 45
            Test.Use #7/30/1912#, 1
            Test.Use #12/24/1926#, 15
            Test.Use #12/25/1926#, 1
            Test.Use #1/7/1989#, 64
            Test.Use #1/8/1989#, 1
            Test.Use #12/31/9999#, 8011
            
        Case "GetEra_WithEraBoundryDate_ReturnsExpected"
            Test.Use #9/8/1868#, 1
            Test.Use #7/29/1912#, 1
            Test.Use #7/30/1912#, 2
            Test.Use #12/24/1926#, 2
            Test.Use #12/25/1926#, 3
            Test.Use #1/7/1989#, 3
            Test.Use #1/8/1989#, 4
            Test.Use #12/31/9999#, 4
            Test.Use #1/1/1868#, 1
            
        Case "GetDaysInMonth_WhenFebuary_ReturnsExpected"
            Test.Use 1, CorLib.Calendar.CurrentEra, 28
            Test.Use 1, 4, 28
            Test.Use 1, 3, 28
            Test.Use 1, 2, 29
            Test.Use 1, 1, 29
            Test.Use 2, 4, 28
            Test.Use 2, 3, 28
            Test.Use 2, 2, 28
            Test.Use 2, 1, 28
            Test.Use 3, 4, 28
            Test.Use 3, 3, 29
            Test.Use 3, 2, 28
            Test.Use 3, 1, 28
            
        Case "GetDaysInMonth_WhenFebuaryInCurrentEra_ReturnsExpected"
            Test.Use 1, 28
            Test.Use 2, 28
            Test.Use 3, 28
            Test.Use 4, 29
            Test.Use 5, 28
            
        Case "GetDaysInYear_WithYearAndEra_ReturnsExpected"
            Test.Use 1, CorLib.Calendar.CurrentEra, 365
            Test.Use 1, 4, 365
            Test.Use 1, 3, 365
            Test.Use 1, 2, 366
            Test.Use 1, 1, 366
            Test.Use 2, CorLib.Calendar.CurrentEra, 365
            Test.Use 2, 4, 365
            Test.Use 2, 3, 365
            Test.Use 2, 2, 365
            Test.Use 2, 1, 365
            Test.Use 4, CorLib.Calendar.CurrentEra, 366
            Test.Use 4, 4, 366
            Test.Use 4, 3, 365
            Test.Use 4, 2, 365
            Test.Use 4, 1, 365
        
        Case "ToDateTime_WithValues_ReturnsExpected"
            Test.Use 1, 9, 8, 1, #9/8/1868 1:01:01 AM#
            Test.Use 45, 7, 29, 1, #7/29/1912 1:01:01 AM#
            Test.Use 1, 7, 30, 2, #7/30/1912 1:01:01 AM#
            Test.Use 15, 12, 24, 2, #12/24/1926 1:01:01 AM#
            Test.Use 1, 12, 25, 3, #12/25/1926 1:01:01 AM#
            Test.Use 64, 1, 7, 3, #1/7/1989 1:01:01 AM#
            Test.Use 1, 1, 8, 4, #1/8/1989 1:01:01 AM#
            Test.Use 8011, 12, 31, 4, #12/31/9999 1:01:01 AM#
            Test.Use 1, 1, 8, Missing, #1/8/1989 1:01:01 AM#
            Test.Use 8011, 12, 31, Missing, #12/31/9999 1:01:01 AM#
            
        Case "ToDateTime_WithValueOutOfSupportedRange_ThrowsArgumentOutOfRangeException"
            Test.Use 1, 9, 7, 1
            Test.Use 8012, 1, 1, 4
            Test.Use 8012, 1, 1, Missing
        
    End Select
End Sub

Public Sub AlgorithmType_WhenCalled_ReturnsSolarCalendar()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.AlgorithmType, Iz.EqualTo(CalendarAlgorithmType.SolarCalendar)
End Sub

Public Sub IsReadOnly_WhenNew_ReturnsFalse()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.IsReadOnly, Iz.False
End Sub

Public Sub Eras_WhenCalled_ReturnsFourEras()
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual() As Long
    Actual = Calendar.Eras
    
    Assert.That Actual, Iz.EqualTo(NewLongs(4, 3, 2, 1))
End Sub

Public Sub TwoDigitYearMax_WhenNew_Returns99()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.TwoDigitYearMax, Iz.EqualTo(99)
End Sub

Public Sub TwoDigitYearMax_WithInvalidValue_ThrowsArgumentOutOfRangeException(ByVal InvalidValue As Long)
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.TwoDigitYearMax = InvalidValue
    
    AssertArgumentOutOfRangeException Err, "Year"
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsDateTimeMaxValue()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.MaxSupportedDateTime, Equals(CorDateTime.MaxValue)
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsStartOfMeijiEra()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.MinSupportedDateTime, Equals(#9/8/1868#)
End Sub

Public Sub AddMethods_WhenCalled_ReturnsExpected()
    Const d As Date = #1/1/1900#
    Dim Calendar As New JapaneseCalendar

    Assert.That Calendar.AddYears(d, 1), Equals(#1/1/1901#), "AddYears #1"
    Assert.That Calendar.AddYears(d, -1), Equals(#1/1/1899#), "AddYears #2"
    Assert.That Calendar.AddMonths(d, 1), Equals(#2/1/1900#), "AddMonths #1"
    Assert.That Calendar.AddMonths(d, -1), Equals(#12/1/1899#), "AddMonths #2"
    Assert.That Calendar.AddDays(d, 1), Equals(#1/2/1900#), "AddDays #1"
    Assert.That Calendar.AddDays(d, -1), Equals(#12/31/1899#)
    Assert.That Calendar.AddHours(d, 1), Equals(#1/1/1900 1:00:00 AM#), "AddHours #1"
    Assert.That Calendar.AddMinutes(d, 1), Equals(#1/1/1900 12:01:00 AM#), "AddMinutes #1"
    Assert.That Calendar.AddSeconds(d, 1), Equals(#1/1/1900 12:00:01 AM#), "AddSeconds #1"
    Assert.That Calendar.AddMilliseconds(d, 1), Equals(NewDateTime(1900, 1, 1, Millisecond:=1)), "AddMilliseconds #1"
    Assert.That Calendar.AddWeeks(d, 5), Equals(#2/5/1900#), "AddWeeks #1"
End Sub

Public Sub AddMilliseconds_WhenResultIsLessThanMinSupported_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.AddMilliseconds Calendar.MinSupportedDateTime, -1
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub AddMilliseconds_WhenResultGreaterThanMaxSupported_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.AddMilliseconds Calendar.MaxSupportedDateTime, 1
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub AddMonths_WhenResultIsLessThanMinSupported_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.AddMonths Calendar.MinSupportedDateTime, -1
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub AddMonths_WhenResultGreaterThanMaxSupported_ThrowsArgumentOutOfRangeException()
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.AddMonths Calendar.MaxSupportedDateTime, 1
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub GetProperties_WhenCalled_ReturnsExpected()
    Dim d As CorDateTime
    Dim Calendar As New JapaneseCalendar
    
    Set d = NewDateTime(2200, 1, 2, 13, 20, 35, 100)
    
    Assert.That Calendar.GetYear(d), Iz.EqualTo(212), "#1"
    Assert.That Calendar.GetMonth(d), Iz.EqualTo(1), "#2"
    Assert.That Calendar.GetDayOfMonth(d), Iz.EqualTo(2), "#3"
    Assert.That Calendar.GetHour(d), Iz.EqualTo(13), "#4"
    Assert.That Calendar.GetMinute(d), Iz.EqualTo(20), "#5"
    Assert.That Calendar.GetSecond(d), Iz.EqualTo(35), "#6"
    Assert.That Calendar.GetMilliseconds(d), Iz.EqualTo(100), "#7"
    Assert.That Calendar.GetDayOfWeek(d), Iz.EqualTo(DayOfWeek.Thursday), "#8"
    Assert.That Calendar.GetDayOfYear(d), Iz.EqualTo(2), "#9"
End Sub

Public Sub GetYear_WithEraBoundryDate_ReturnsExpected(ByVal Value As Date, ByVal Expected As Long)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetYear(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetEra_WithEraBoundryDate_ReturnsExpected(ByVal Value As Date, ByVal Expected As Long)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetEra(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_WhenFebuary_ReturnsExpected(ByVal Year As Long, ByVal Era As Long, ByVal Expected As Long)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(Year, 2, Era)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInMonth_WhenFebuaryInCurrentEra_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInMonth(Year, 2)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDaysInYear_WithYearAndEra_ReturnsExpected(ByVal Year As Long, ByVal Era As Long, ByVal Expected As Long)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDaysInYear(Year, Era)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToDateTime_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByRef Era As Variant, ByVal Expected As Date)
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.ToDateTime(Year, Month, Day, 1, 1, 1, 0, Era)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub ToDateTime_WithValueOutOfSupportedRange_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByRef Era As Variant)
    Dim Calendar As New JapaneseCalendar
    
    On Error Resume Next
    Calendar.ToDateTime Year, Month, Day, 1, 1, 1, 1, Era
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub IsReadOnly_WhenMadeReadOnly_ReturnsTrue()
    Dim cal As New JapaneseCalendar
    
    Dim Actual As Calendar
    Set Actual = CorLib.Calendar.ReadOnly(cal)
    
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub Clone_WhenCalled_ReturnsWritableClone()
    Dim Calendar As New JapaneseCalendar
    Calendar.TwoDigitYearMax = 5400
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Dim Actual As JapaneseCalendar
    Set Actual = Calendar.Clone
    
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(5400)
    Assert.That Actual.IsReadOnly, Iz.False
End Sub

Public Sub TestSerialization()
    Dim Calendar As New JapaneseCalendar
    Calendar.TwoDigitYearMax = 5400
    Set Calendar = CorLib.Calendar.ReadOnly(Calendar)
    
    Dim Actual As Calendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.False
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(5400)
End Sub

Public Sub IsLeapMonth_WhenCalled_AlwaysReturnsFalse()
    Dim Calendar As New JapaneseCalendar
    
    Assert.That Calendar.IsLeapMonth(1, 2), Iz.False
End Sub

Public Sub GetLeapMonth_WhenCalled_ReturnsZero()
    Dim Calendar As New JapaneseCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetLeapMonth(2004)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

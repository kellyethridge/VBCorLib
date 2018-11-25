VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HijriCalendarTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HijriCalendarTests
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource


Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "HijriAdjustment_WithOutOfRangeValues_ThrowsArgumentOutOfRangeException"
            Test.Use -3
            Test.Use 3
    
        Case "GetMonth_WithValue_ReturnsExpected"
            Test.Use #7/2/2017#, 10
            Test.Use #1/1/2017#, 4
            Test.Use #1/27/2017#, 4
            Test.Use #1/28/2017#, 5
            Test.Use #2/26/2017#, 5
            Test.Use #2/27/2017#, 6
            Test.Use #8/7/650#, 12
            Test.Use #9/5/650#, 12
            Test.Use #9/22/2082#, 12
            
        Case "ToDateTime_WithValues_ReturnsExpected"
            Test.Use 1, 1, 1, #7/18/622#
            Test.Use 2, 1, 1, #7/7/623#
            Test.Use 2017, 12, 30, #6/25/2579#
            
        Case "ToDateTime_WithHijriAdjustment_ReturnsExpected"
            Test.Use -1, 2017, 12, 30, #6/26/2579#
            Test.Use -2, 2, 6, 29, #1/1/624#
            Test.Use -1, 2, 6, 29, #12/31/623#
            
        Case "ToDateTime_WithInvalidDay_ThrowsArgumentOutOfRangeException"
            Test.Use 29, 1, 0
            Test.Use 29, 1, 31
            Test.Use 29, 12, 31
            Test.Use 28, 12, 30
            
        Case "GetDayOfMonth_WithValue_ReturnsExpected"
            Test.Use #1/16/1000#, 1
            Test.Use #2/13/1000#, 29
            Test.Use #2/14/1000#, 1
            Test.Use #3/15/1000#, 30
            Test.Use #3/16/1000#, 1
            Test.Use #4/13/1000#, 29
            Test.Use #4/14/1000#, 1
            Test.Use #5/13/1000#, 30
            Test.Use #5/14/1000#, 1
            Test.Use #6/11/1000#, 29
            Test.Use #6/12/1000#, 1
            Test.Use #7/11/1000#, 30
            Test.Use #7/12/1000#, 1
            Test.Use #8/9/1000#, 29
            Test.Use #8/10/1000#, 1
            Test.Use #9/8/1000#, 30
            Test.Use #9/9/1000#, 1
            Test.Use #10/7/1000#, 29
            Test.Use #10/8/1000#, 1
            Test.Use #11/6/1000#, 30
            Test.Use #11/7/1000#, 1
            Test.Use #12/5/1000#, 29
            Test.Use #12/6/1000#, 1
            Test.Use #1/4/1001#, 30
            Test.Use #6/25/624#, 30
            
        Case "GetDayOfYear_WithValue_ReturnsExpected"
            Test.Use #12/6/1000#, 1
            Test.Use #11/24/1001#, 354
            Test.Use #11/25/1001#, 1
            Test.Use #7/7/623#, 1
            Test.Use #6/25/624#, 355
            Test.Use #6/26/624#, 1
            
        Case "GetWeekOfYear_WithFirstDayRule_ReturnsExpected"
            Test.Use #5/25/627#, DayOfWeek.Friday, 1
            Test.Use #5/12/628#, DayOfWeek.Friday, 51
            Test.Use #5/13/628#, DayOfWeek.Friday, 1
            
        Case "ToFourDigitYear_WithValue_ReturnsExpected"
            Test.Use 100, 100
            Test.Use 50, 1450
            
    End Select
End Sub

Public Sub AlgorithmType_WhenCalled_ReturnsLunarCalendar()
    Dim Calendar As New HijriCalendar
    
    Assert.That Calendar.AlgorithmType, Iz.EqualTo(CalendarAlgorithmType.LunarCalendar)
End Sub

Public Sub MinSupportedDateTime_WhenCalled_ReturnsDateTimeMinValue()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MinSupportedDateTime
    
    Assert.That Actual, Equals(#7/18/622#)
End Sub

Public Sub MaxSupportedDateTime_WhenCalled_ReturnsDateTimeMaxValue()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.MaxSupportedDateTime
    
    Assert.That Actual, Equals(CorDateTime.MaxValue)
End Sub

Public Sub HijriAdjustment_WithOutOfRangeValues_ThrowsArgumentOutOfRangeException(ByVal Value As Long)
    Dim Calendar As New HijriCalendar
    
    On Error Resume Next
    Calendar.HijriAdjustment = Value
    
    AssertArgumentOutOfRangeException Err, "HijriAdjustment", Value
End Sub

Public Sub ToDateTime_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As Date)
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.ToDateTime(Year, Month, Day, 1, 1, 1, 1).DateOnly
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub ToDateTime_WithHijriAdjustment_ReturnsExpected(ByVal Adjustment As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Expected As Date)
    Dim Calendar    As New HijriCalendar
    Dim Actual      As CorDateTime
    
    Calendar.HijriAdjustment = Adjustment
    Set Actual = Calendar.ToDateTime(Year, Month, Day, 1, 1, 1, 1).DateOnly
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub ToDateTime_WithInvalidDay_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long)
    Dim Calendar As New HijriCalendar
    
    On Error Resume Next
    Calendar.ToDateTime Year, Month, Day, 1, 1, 1, 1
    
    AssertArgumentOutOfRangeException Err, "Day"
End Sub

Public Sub IsLeapYear_WhenCalled_ReturnsExpected()
    Dim Calendar As New HijriCalendar
    
    Assert.That Calendar.IsLeapYear(1), Iz.False, "#1"
    Assert.That Calendar.IsLeapYear(2), Iz.True, "#2"
    Assert.That Calendar.IsLeapYear(3), Iz.False, "#3"
    Assert.That Calendar.IsLeapYear(4), Iz.False, "#4"
    Assert.That Calendar.IsLeapYear(5), Iz.True, "#5"
    Assert.That Calendar.IsLeapYear(6), Iz.False, "#6"
    Assert.That Calendar.IsLeapYear(7), Iz.True, "#7"
    Assert.That Calendar.IsLeapYear(8), Iz.False, "#8"
    Assert.That Calendar.IsLeapYear(9), Iz.False, "#9"
    Assert.That Calendar.IsLeapYear(10), Iz.True, "#10"
    Assert.That Calendar.IsLeapYear(11), Iz.False, "#11"
    Assert.That Calendar.IsLeapYear(12), Iz.False, "#12"
    Assert.That Calendar.IsLeapYear(13), Iz.True, "#13"
    Assert.That Calendar.IsLeapYear(14), Iz.False, "#14"
    Assert.That Calendar.IsLeapYear(15), Iz.False, "#15"
    Assert.That Calendar.IsLeapYear(16), Iz.True, "#16"
    Assert.That Calendar.IsLeapYear(17), Iz.False, "#17"
    Assert.That Calendar.IsLeapYear(18), Iz.True, "#18"
    Assert.That Calendar.IsLeapYear(19), Iz.False, "#19"
    Assert.That Calendar.IsLeapYear(20), Iz.False, "#20"
    Assert.That Calendar.IsLeapYear(21), Iz.True, "#21"
    Assert.That Calendar.IsLeapYear(22), Iz.False, "#22"
    Assert.That Calendar.IsLeapYear(23), Iz.False, "#23"
    Assert.That Calendar.IsLeapYear(24), Iz.True, "#24"
    Assert.That Calendar.IsLeapYear(25), Iz.False, "#25"
    Assert.That Calendar.IsLeapYear(26), Iz.True, "#26"
    Assert.That Calendar.IsLeapYear(27), Iz.False, "#27"
    Assert.That Calendar.IsLeapYear(28), Iz.False, "#28"
    Assert.That Calendar.IsLeapYear(29), Iz.True, "#29"
    Assert.That Calendar.IsLeapYear(30), Iz.False, "#30"
    Assert.That Calendar.IsLeapYear(31), Iz.False, "#31"
    Assert.That Calendar.IsLeapYear(32), Iz.True, "#32"
End Sub

Public Sub GetYear_WhenCalled_ReturnsExpected()
    Dim Calendar As New HijriCalendar
    
    Assert.That Calendar.GetYear(#7/18/622#), Iz.EqualTo(1)
    Assert.That Calendar.GetYear(#6/11/2017#), Iz.EqualTo(1438)
End Sub

Public Sub GetMonth_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetMonth(d)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDayOfMonth_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Calendar As New HijriCalendar

    Dim Actual As Long
    Actual = Calendar.GetDayOfMonth(d)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetDayOfYear_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetDayOfYear(d)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub AddYears_WithValue_ReturnsExpected()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.AddYears(#1/1/2000#, 1)
    
    Assert.That Actual, Equals(#12/21/2000#)
End Sub

Public Sub AddMonths_WithValue_ReturnsExpected()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.AddMonths(#1/1/1000#, 1)
    
    Assert.That Actual, Equals(#1/31/1000#)
End Sub

Public Sub AddDays_WithValue_ReturnsExpected()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As CorDateTime
    Set Actual = Calendar.AddDays(#1/1/2000#, 1000)
    
    Assert.That Actual, Equals(#9/27/2002#)
End Sub

Public Sub GetWeekOfYear_WithFirstDayRule_ReturnsExpected(ByVal d As Date, ByVal FirstDay As DayOfWeek, ByVal Expected As Long)
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstDay, FirstDay)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestSerialization()
    Dim Calendar As New HijriCalendar
    Calendar.HijriAdjustment = 1
    Calendar.TwoDigitYearMax = 1000
    
    Dim Actual As HijriCalendar
    Set Actual = Persist(Calendar)
    
    Assert.That Actual.HijriAdjustment, Iz.EqualTo(1)
    Assert.That Actual.TwoDigitYearMax, Iz.EqualTo(1000)
End Sub

Public Sub IsReadOnly_WhenNew_ReturnsFalse()
    Dim Calendar As New HijriCalendar
    
    Assert.That Calendar.IsReadOnly, Iz.False
End Sub

Public Sub ReadOnly_WhenCalled_ReturnsReadOnlyInstance()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Calendar
    Set Actual = CorLib.Calendar.ReadOnly(Calendar)
    
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub AddMethods_WhenCalled_ReturnsExpected()
    Const d As Date = #1/1/1800#
    Dim Calendar As New HijriCalendar

    Assert.That Calendar.AddYears(d, 1), Equals(#12/21/1800#), "#1"
    Assert.That Calendar.AddMonths(d, 1), Equals(#1/30/1800#), "#2"
    Assert.That Calendar.AddMonths(d, -1), Equals(#12/2/1799#), "#3"
    Assert.That Calendar.AddDays(d, 1), Equals(#1/2/1800#), "#4"
    Assert.That Calendar.AddHours(d, 1), Equals(#1/1/1800 1:00:00 AM#), "#5"
    Assert.That Calendar.AddMinutes(d, 1), Equals(#1/1/1800 12:01:00 AM#), "#6"
    Assert.That Calendar.AddSeconds(d, 1), Equals(#1/1/1800 12:00:01 AM#), "#7"
    Assert.That Calendar.AddMilliseconds(d, 1), Equals(NewDateTime(1800, 1, 1, Millisecond:=1)), "#8"
    Assert.That Calendar.AddWeeks(d, 5), Equals(#2/5/1800#), "#9"
End Sub

Public Sub GetProperties_WhenCalled_ReturnsExpected()
    Const d As Date = #1/2/2200 1:20:35 PM#
    Dim Calendar As New HijriCalendar
    
    Assert.That Calendar.GetYear(d), Iz.EqualTo(1626), "#1"
    Assert.That Calendar.GetMonth(d), Iz.EqualTo(11), "#2"
    Assert.That Calendar.GetDayOfMonth(d), Iz.EqualTo(16), "#3"
    Assert.That Calendar.GetHour(d), Iz.EqualTo(13), "#4"
    Assert.That Calendar.GetMinute(d), Iz.EqualTo(20), "#5"
    Assert.That Calendar.GetSecond(d), Iz.EqualTo(35), "#6"
    Assert.That Calendar.GetDayOfWeek(d), Iz.EqualTo(DayOfWeek.Thursday), "#7"
    Assert.That Calendar.GetDayOfYear(d), Iz.EqualTo(311), "#8"
End Sub

Public Sub GetMilliseconds_WhenCalled_ReturnsMillisecondsPart()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetMilliseconds(NewDateTime(Second:=1, Millisecond:=50))
    
    Assert.That Actual, Iz.EqualTo(50)
End Sub

Public Sub ToFourDigitYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Long)
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.ToFourDigitYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TwoDigitYearMax_WhenNew_Returns1451()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.TwoDigitYearMax
    
    Assert.That Actual, Iz.EqualTo(1451)
End Sub

Public Sub GetLeapMonth_WhenCalled_ReturnsZero()
    Dim Calendar As New HijriCalendar
    
    Dim Actual As Long
    Actual = Calendar.GetLeapMonth(2)
    
    Assert.That Actual, Iz.EqualTo(0)
End Sub

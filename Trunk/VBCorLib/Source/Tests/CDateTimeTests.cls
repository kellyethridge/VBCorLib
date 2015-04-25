VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDateTimeTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDateTime
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToOADate_WithValues_ReturnsExpected"
            Test.Use #1/1/2004 8:30:30 AM#
            Test.Use #1/1/1503 8:30:30 AM#
            Test.Use #1/1/100#
            Test.Use #1/1/9999 11:59:59 PM#
            
        Case "FromOADate_WithValues_ReturnscDateTime"
            Test.Use #12:00:00 AM#
            Test.Use #5/5/1900 3:30:00 PM#
            Test.Use #1/1/100#
            Test.Use #1/1/9999 11:59:59 PM#
            
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use #1/2/2004 3:04:05 AM#, True
            Test.Use #1/2/2004 3:04:06 AM#, False
            Test.Use NewDateTime(2004, 1, 2, 3, 4, 5), True
            Test.Use NewDateTime(2004, 1, 2, 3, 4, 6), False
            Test.Use Null, False
            Test.Use Nothing, False
            Test.Use Empty, False
            
        Case "Year_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 1899
            Test.Use #1/1/100#, 100
            Test.Use #1/1/9999#, 9999
            
        Case "Month_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 12
            Test.Use #1/1/100#, 1
            Test.Use #12/31/9999 11:59:59 PM#, 12
            
        Case "Day_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 30
            Test.Use #1/1/100#, 1
            Test.Use #12/31/9999 11:59:59 PM#, 31
            Test.Use #2/29/2016#, 29
            
        Case "Hour_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 0
            Test.Use #1/1/100#, 0
            Test.Use #12/31/9999 11:59:59 PM#, 23
        
        Case "Minute_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 0
            Test.Use #12:59:00 AM#, 59
            Test.Use #1/1/100#, 0
            Test.Use #12/31/9999 11:59:59 PM#, 59
            
        Case "Second_WithValue_ReturnsExpected"
            Test.Use #12:00:00 AM#, 0
            Test.Use #12:00:59 AM#, 59
            Test.Use #1/1/100#, 0
            Test.Use #12/31/9999 11:59:59 PM#, 59
        
        Case "IsLeapYear_WithValue_ReturnsExpected"
            Test.Use 2015, False
            Test.Use 2016, True
            Test.Use 2017, False
            
        Case "DayOfYear_WithValue_ReturnsExpected"
            Test.Use #1/1/100#, 1
            Test.Use #12/31/9999#, 365
            Test.Use #3/1/2015#, 60
            Test.Use #2/29/2016#, 60
            Test.Use #12/31/2016#, 366
            
        Case "DayOfWeek_WithValue_ReturnsExpected"
            Test.Use #1/1/100#, DayOfWeek.Friday
            Test.Use #1/2/100#, DayOfWeek.Saturday
            Test.Use #1/3/100#, DayOfWeek.Sunday
            Test.Use #1/4/100#, DayOfWeek.Monday
            Test.Use #1/5/100#, DayOfWeek.Tuesday
            Test.Use #1/6/100#, DayOfWeek.Wednesday
            Test.Use #1/7/100#, DayOfWeek.Thursday
            Test.Use #1/8/100#, DayOfWeek.Friday
            
        Case "AddMonths_WithValues_ReturnsExpected"
            Test.Use 1, NewDateTime(2001, 2, 1)
            Test.Use 2, NewDateTime(2001, 3, 1)
            Test.Use 3, NewDateTime(2001, 4, 1)
            Test.Use 4, NewDateTime(2001, 5, 1)
            Test.Use 5, NewDateTime(2001, 6, 1)
            Test.Use 6, NewDateTime(2001, 7, 1)
            Test.Use 7, NewDateTime(2001, 8, 1)
            Test.Use 8, NewDateTime(2001, 9, 1)
            Test.Use 9, NewDateTime(2001, 10, 1)
            Test.Use 10, NewDateTime(2001, 11, 1)
            Test.Use 11, NewDateTime(2001, 12, 1)
            Test.Use 12, NewDateTime(2002, 1, 1)
            
        Case "DaysInMonth_WithValues_ReturnsExpected"
            Test.Use 2014, 1, 31
            Test.Use 2014, 2, 28
            Test.Use 2014, 3, 31
            Test.Use 2014, 4, 30
            Test.Use 2014, 5, 31
            Test.Use 2014, 6, 30
            Test.Use 2014, 7, 31
            Test.Use 2014, 8, 31
            Test.Use 2014, 9, 30
            Test.Use 2014, 10, 31
            Test.Use 2014, 11, 30
            Test.Use 2014, 12, 31
            Test.Use 2016, 2, 29
            
        Case "CompareTo_WithValues_ReturnsExpected"
            Test.Use 5, 10, -1
            Test.Use 10, 5, 1
            Test.Use 10, 10, 0
            
        Case "Subtract_WithTimeSpans_ReturnsExpected"
            Test.Use TimeSpan.Zero, NewDateTime(2001, 1, 1)
            Test.Use TimeSpan.FromHours(-1), NewDate(#1/1/2001 1:00:00 AM#)
            Test.Use TimeSpan.FromHours(1), NewDate(#12/31/2000 11:00:00 PM#)
            
        Case "Subtract_WithDates_ReturnsExpected"
            Test.Use #12/31/2000#, TimeSpan.FromDays(1)
            Test.Use NewDateTime(2000, 12, 31), TimeSpan.FromDays(1)
            
        Case "DaysInMonth_WithInvalidValues_ThrowsArgumentOutOfRangeException"
            Test.Use 0, 1, "Year"
            Test.Use 10000, 1, "Year"
            Test.Use 1, 0, "Month"
            Test.Use 1, 13, "Month"
            
    End Select
End Sub

Public Sub ToOADate_WithDefaultDate_ReturnsMidnight()
    Dim Source As New cDateTime
    
    Dim Actual As Date
    Actual = Source.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#12:00:00 AM#)
End Sub

Public Sub ToOADate_WithValues_ReturnsExpected(ByVal Expected As Date)
    Dim Source As cDateTime
    Set Source = NewDateTime(Year(Expected), Month(Expected), Day(Expected), Hour(Expected), Minute(Expected), Second(Expected))
    
    Dim Actual As Date
    Actual = Source.ToOADate
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromOADate_WithValues_ReturnscDateTime(ByVal Expected As Date)
    Dim dt As cDateTime
        
    Set dt = cDateTime.FromOADate(Expected)
    
    Dim Actual As Date
    Actual = dt.ToOADate
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByVal Other As Variant, ByVal Expected As Boolean)
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 1, 2, 3, 4, 5)
    
    Dim Actual As Boolean
    Actual = dt.Equals(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Year_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Year
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Month_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Month
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Day_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Day
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Hour_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Hour
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Minute_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Minute
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Second_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.Second
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Millisecond_WhenCalled_ReturnsMillisecondComponent()
    Dim dt As cDateTime
    Set dt = NewDateTime(1, 1, 1, 1, 1, 1, 37)
    
    Dim Actual As Long
    Actual = dt.Millisecond
    
    Assert.That Actual, Iz.EqualTo(37)
End Sub

Public Sub IsLeapYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = cDateTime.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Constructor_WithTimeComponentsOnly_SetsToOADate()
    Dim dt As cDateTime
    
    Set dt = NewDateTime(Hour:=18, Minute:=39, Second:=52)
    
    Assert.That dt.ToOADate, Iz.EqualTo(#6:39:52 PM#)
End Sub

Public Sub DayOfYear_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As Long
    Actual = dt.DayOfYear
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DayOfWeek_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As DayOfWeek)
    Dim dt As cDateTime
    Set dt = cDateTime.FromOADate(d)
    
    Dim Actual As DayOfWeek
    Actual = dt.DayOfWeek
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TotalMilliseconds_WhenCalled_ReturnsExpected()
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 1, 1, 1, 1, 1, 1)
    
    Dim Actual As Currency
    Actual = dt.TotalMilliseconds
    
    Const Expected As Currency = 63208515661001@
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestTicks()
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 1, 1, 1, 1, 1, 1)
    
    Dim Actual As Variant
    Actual = dt.Ticks
    
    Assert.That Actual, Iz.EqualTo(CDec("632085156610010000"))
End Sub

Public Sub DateOnly_WhenCalled_ReturnsDateTimeWithDatePartOnly()
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 1, 1, 1, 1, 1, 1)
    
    Set dt = dt.DateOnly
    
    Dim Actual As Boolean
    Actual = dt.Equals(#1/1/2004#)
    Assert.That Actual, Iz.True
End Sub

Public Sub TimeOfDay_WhenCalled_ReturnsExpected()
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 1, 1, 6, 10, 15, 100)
    
    Dim Actual As TimeSpan
    Set Actual = dt.TimeOfDay
    
    Dim Expected As TimeSpan
    Set Expected = newtimespan(Hours:=6, Minutes:=10, Seconds:=15, Milliseconds:=100)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Add_WithOneHour_ReturnsDateTimePlus1Hour()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    Dim ts As TimeSpan
    Set ts = TimeSpan.FromHours(1)
    
    Dim Actual As cDateTime
    Set Actual = dt.Add(ts)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 1, 1)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMilliseconds_With100_ReturnsDateTimePlus100Milliseconds()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddMilliseconds(100)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 1, Millisecond:=100)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddDays_With5_ReturnsDateTimePlus5Days()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddDays(5)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 6)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddHours_With5_ReturnsDateTimePlus5Hours()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddHours(5)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 1, 5)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMinutes_With5_ReturnsDateTimePlus5Minutes()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddMinutes(5)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 1, Minute:=5)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddSeconds_With60_ReturnsDateTimePlus1Minute()
    Dim dt As cDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddSeconds(60)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2000, 1, 1, Minute:=1)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMonths_WithValues_ReturnsExpected(ByVal Months As Long, ByVal Expected As cDateTime)
    Dim dt As cDateTime
    Set dt = NewDateTime(2001, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddMonths(Months)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddYear_With5_ReturnsDateTimePlus5Years()
    Dim dt As cDateTime
    Set dt = NewDateTime(2004, 5, 16)
    
    Dim Actual As cDateTime
    Set Actual = dt.AddYears(5)
    
    Dim Expected As cDateTime
    Set Expected = NewDateTime(2009, 5, 16)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub DaysInMonth_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = cDateTime.DaysInMonth(Year, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DaysInMonth_WithInvalidValues_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal ParameterName As String)
    On Error GoTo errTrap
    
    cDateTime.DaysInMonth Year, Month
    
errTrap:
    AssertArgumentOutOfRangeException Err, ParameterName
End Sub

Public Sub CompareTo_WithValues_ReturnsExpected(ByVal LeftHours As Double, ByVal RightHours As Double, ByVal Expected As Long)
    Dim Left As cDateTime
    Dim Right As cDateTime
    Set Left = cDateTime.MinValue.AddHours(LeftHours)
    Set Right = cDateTime.MinValue.AddHours(RightHours)
    
    Dim Actual As Long
    Actual = Left.CompareTo(Right)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Subtract_WithTimeSpans_ReturnsExpected(ByVal TimeSpanToSubtract As TimeSpan, ByVal Expected As cDateTime)
    Dim dt As cDateTime
    Set dt = NewDateTime(2001, 1, 1)
    
    Dim Actual As cDateTime
    Set Actual = dt.Subtract(TimeSpanToSubtract)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Subtract_WithDates_ReturnsExpected(ByVal DateToSubtract As Variant, ByVal Expected As TimeSpan)
    Dim dt As cDateTime
    Set dt = NewDate(#1/1/2001#)
    
    Dim Actual As TimeSpan
    Set Actual = dt.Subtract(DateToSubtract)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Subtract_WithNothing_ThrowsArgumentNullException()
    On Error GoTo errTrap
    Dim dt As cDateTime
    Set dt = NewDate(#1/1/2001#)
    
    dt.Subtract Nothing
    
errTrap:
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub Subtract_WithNonsupportedObject_ThrowsArgumentException()
    On Error GoTo errTrap
    Dim dt As cDateTime
    Set dt = NewDate(#1/1/2001#)
    
    Dim Value As New Collection
    dt.Subtract Value
    
errTrap:
    AssertArgumentException Err, "Value"
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToFileTime()
    Dim dt As cDateTime
    
    Set dt = NewDateTime(2004, 6, 6, 6, 6, 6, 6)
    Assert.AreEqual 12731000766006@, dt.ToFileTime
    
    Set dt = NewDateTime(2004, 1, 6, 6, 6, 6, 6)
    Assert.AreEqual 12717871566006@, dt.ToFileTime
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToFileTimeUtc()
    Dim dt As cDateTime
    
    Set dt = NewDateTime(2004, 6, 6, 6, 6, 6, 6)
    Assert.AreEqual 12730975566006@, dt.ToFileTimeUtc
    
    Set dt = NewDateTime(2004, 1, 6, 6, 6, 6, 6)
    Assert.AreEqual 12717842766006@, dt.ToFileTimeUtc
End Sub

Public Sub Now_WhenCalled_ReturnsLocalTime()
    Dim dt As cDateTime

    Set dt = cDateTime.Now
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.LocalKind)
End Sub

Public Sub Now_WhenCalled_ReturnsCurrentDateTime()
    Dim Actual As Long
    
    Actual = cDateTime.Now.Subtract(NewDateTime(1970, 1, 1)).TotalSeconds
    
    Dim Diff As Long
    Diff = DateDiff("s", #1/1/1970#, Now) - Actual
    Assert.That Diff, Iz.GreaterThanOrEqualTo(-5).AndAlso.LessThanOrEqualTo(5)
End Sub

Public Sub UtcNow_WhenCalled_ReturnsUtcTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.UtcNow
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.UtcKind)
End Sub

Public Sub Today_WhenCalled_ReturnsLocalTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.Today
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.LocalKind)
End Sub

Public Sub Today_WhenCalled_ReturnsCurrentDate()
    Dim Actual As cDateTime
    
    Set Actual = cDateTime.Today
    
    Assert.That Actual.Equals(Date), Iz.True
End Sub

Public Sub Parse_WithDate_ReturnsParsedDateTime()
    Dim Actual As cDateTime
    
    Set Actual = cDateTime.Parse("1/1/2004 8:30:30am")
    
    Assert.That Actual, Equals(NewDate(#1/1/2004 8:30:30 AM#))
End Sub

Public Sub FromFileTimeUtc_WhenCalled_ReturnsExpected()
    Dim Actual As cDateTime
    
    Set Actual = cDateTime.FromFileTimeUtc(12.3456)
    
    Assert.That Actual.Kind, Iz.EqualTo(DateTimeKind.UtcKind)
    Assert.That Actual.Ticks, Iz.EqualTo(CDec("504911232000123456"))
End Sub

Public Sub FromFileTimeUtc_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error GoTo errTrap
    
    cDateTime.FromFileTimeUtc -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Time"
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToLocalTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/2004 12:00:00 PM#)
    Set dt = dt.ToLocalTime
    Assert.IsTrue dt.Equals(#1/1/2004 4:00:00 AM#), "ToLocalTime should return 1/1/2004 4AM."
    
    Set dt = cDateTime.FromOADate(#6/1/2004 12:00:00 PM#)
    Set dt = dt.ToLocalTime
    Assert.IsTrue dt.Equals(#6/1/2004 5:00:00 AM#), "ToLocalTime should return 6/1/2004 5AM."
End Sub

' This test is subject to the current timezone. Modify this test
' as necessary to your timezone.
Public Sub TestToUniversalTime()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/2004 12:00:00 PM#)
    Set dt = dt.ToUniversalTime
    Assert.IsTrue dt.Equals(#1/1/2004 8:00:00 PM#), "ToUniversalTime should return 1/1/2004 8PM."
    
    Set dt = cDateTime.FromOADate(#6/1/2004 12:00:00 PM#)
    Set dt = dt.ToUniversalTime
    Assert.IsTrue dt.Equals(#6/1/2004 7:00:00 PM#), "ToUniversalTime should return 6/1/2004 7PM."
End Sub

Public Sub TestQuotedFormatting()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/5/2005#)
    Assert.AreEqual "5 tt 05", dt.ToString("d 'tt' 05")
End Sub

Public Sub TestFromBinary()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromBinary(-915885960650477.5808@)
    Assert.AreEqual DateTimeKind.LocalKind, dt.Kind
    Assert.AreEqual 205, dt.Year
    Assert.AreEqual 6, dt.Month
    Assert.AreEqual 7, dt.Day
    Assert.AreEqual 20, dt.Hour
    Assert.AreEqual 57, dt.Minute
    Assert.AreEqual 15, dt.Second
End Sub
'205, 6, 7, 20, 57, 15

Public Sub TestToBinary()
    Dim dt As cDateTime
    
    Set dt = NewDate(#6/7/205 8:57:15 PM#, LocalKind)
    Assert.AreEqual -915885960650477.5808@, dt.ToBinary
End Sub

Public Sub TestMaxDateToString()
    Dim dt As cDateTime
    
    Set dt = NewDate(#12/31/9999 11:59:59 PM#)
    Assert.AreEqual "12/31/9999 11:59:59 PM", dt.ToString("MM/dd/yyyy hh:mm:ss tt")
End Sub


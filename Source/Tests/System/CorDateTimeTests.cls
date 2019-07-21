VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorDateTimeTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CorDateTimeTests
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
            
        Case "FromOADate_WithValue_ReturnsDateTime"
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
            Test.Use 3, False
            
        Case "Equals_WithMultipleArguments_ReturnsExpected"
            Test.Use #1/2/2001#, True
            Test.Use #1/1/2001#, False
            Test.Use NewDateTime(2001, 1, 2), True
            Test.Use 3, False
                   
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
            Test.Use #1/1/2001#, 0
            Test.Use NewDate(#1/1/2001#), 0
            Test.Use #1/1/2000#, 1
            Test.Use NewDate(#1/1/2000#), 1
            Test.Use #1/1/2002#, -1
            Test.Use NewDate(#1/1/2002#), -1
            
        Case "Subtract_WithTimeSpans_ReturnsExpected"
            Test.Use TimeSpan.Zero, NewDateTime(2001, 1, 1)
            Test.Use TimeSpan.FromHours(-1), NewDate(#1/1/2001 1:00:00 AM#)
            Test.Use TimeSpan.FromHours(1), NewDate(#12/31/2000 11:00:00 PM#)
            
        Case "Subtract_WithDate_ReturnsExpected"
            Test.Use #12/31/2000#, TimeSpan.FromDays(1)
            Test.Use NewDateTime(2000, 12, 31), TimeSpan.FromDays(1)
            
        Case "DaysInMonth_WithInvalidValues_ThrowsArgumentOutOfRangeException"
            Test.Use 0, 1, "Year"
            Test.Use 10000, 1, "Year"
            Test.Use 1, 0, "Month"
            Test.Use 1, 13, "Month"
            
        Case "ToFileTime_WithKinds_ReturnsExpected"
            Test.Use DateTimeKind.UnspecifiedKind, 13079206800000@
            Test.Use DateTimeKind.LocalKind, 13079206800000@
            Test.Use DateTimeKind.UtcKind, 13079181600000@
            
        Case "ToFileTimeUtc_WithKinds_ReturnsExpected"
            Test.Use DateTimeKind.LocalKind, 13079206800000@
            Test.Use DateTimeKind.UnspecifiedKind, 13079181600000@
            Test.Use DateTimeKind.UtcKind, 13079181600000@

        Case "ToUniversalTime_WithKinds_ReturnsUtcKind"
            Test.Use DateTimeKind.UtcKind
            Test.Use DateTimeKind.UnspecifiedKind
            Test.Use DateTimeKind.LocalKind
            
        Case "ToLocalTime_WithKinds_ReturnsLocalKind"
            Test.Use DateTimeKind.UtcKind
            Test.Use DateTimeKind.UnspecifiedKind
            Test.Use DateTimeKind.LocalKind
        
        Case "ToString_WithCultureInfo_ReturnsExpected"
            Test.Use "d", "en-US", "6/15/2009"
            Test.Use "d", "fr-FR", "15/06/2009"
            Test.Use "d", "ja-JP", "2009/06/15"
            
        Case "ToString_WithFormat_ReturnsExpected"
            Test.Use "", "02/03/2001 13:05:06"
            Test.Use "d", "02/03/2001"
            Test.Use "dddd ddd dd d", "Saturday Sat 03 3"
            Test.Use "yyyy yy", "2001 01"
            Test.Use "MMMM MMM MM M", "February Feb 02 2"
            Test.Use "HH H", "13 13"
            Test.Use "hh h", "01 1"
            Test.Use "mm m", "05 5"
            Test.Use "ss s", "06 6"
            Test.Use "fff ff f", "123 12 1"
            Test.Use "tt t", "PM P"
            Test.Use "HH:mm:ss", "13:05:06"
            Test.Use "h:mm:ss tt", "1:05:06 PM"
            Test.Use "\d", "d"
            Test.Use "\dddd", "dSat"
            Test.Use "%z", "-8"
            Test.Use "%zz", "-08"
            Test.Use "%zzz", "-08:00"
            Test.Use "%zzzz", "-08:00"
            Test.Use "'HH H'", "HH H"
            Test.Use "'\'HH H\''", "'HH H'"
            
        Case "ToString_WithFractionsOfSecond_ReturnsExpected"
            Test.Use "%f", "7"
            Test.Use "ff", "76"
            Test.Use "fff", "765"
            Test.Use "ffff", "7654"
            Test.Use "fffff", "76543"
            Test.Use "ffffff", "765432"
            Test.Use "fffffff", "7654321"
            
        Case "ToString_WithValue_ReturnsExpected"
            Test.Use CorDateTime.MinValue, "1/1/0001 12:00:00 AM"
            Test.Use CorDateTime.MaxValue, "12/31/9999 11:59:59 PM"
            
    End Select
End Sub

Public Sub NewDateTime_WhenCalled_CreatesInstance()
    Dim Actual As CorDateTime
    
    Set Actual = NewDateTime(1000, 2, 3, 4, 5, 6, 7)
    
    Assert.That Actual.Year, Iz.EqualTo(1000)
    Assert.That Actual.Month, Iz.EqualTo(2)
    Assert.That Actual.Day, Iz.EqualTo(3)
    Assert.That Actual.Hour, Iz.EqualTo(4)
    Assert.That Actual.Minute, Iz.EqualTo(5)
    Assert.That Actual.Second, Iz.EqualTo(6)
    Assert.That Actual.Millisecond, Iz.EqualTo(7)
End Sub

Public Sub NewDate_WhenCalled_CreatesInstance()
    Dim Actual As CorDateTime
    
    Set Actual = NewDate(#2/3/1000 4:05:06 AM#)
    
    Assert.That Actual.Year, Iz.EqualTo(1000)
    Assert.That Actual.Month, Iz.EqualTo(2)
    Assert.That Actual.Day, Iz.EqualTo(3)
    Assert.That Actual.Hour, Iz.EqualTo(4)
    Assert.That Actual.Minute, Iz.EqualTo(5)
    Assert.That Actual.Second, Iz.EqualTo(6)
    Assert.That Actual.Millisecond, Iz.EqualTo(0)
End Sub

Public Sub SpecifyKind_WhenSameAsCurrentValue_ReturnsPassedInInstance()
    Dim dt As CorDateTime
    Set dt = NewDate(#1/1/2000#, LocalKind)
    
    Dim Actual As CorDateTime
    Set Actual = CorDateTime.SpecifyKind(dt, LocalKind)
    
    Assert.That Actual, Iz.SameAs(dt)
End Sub

Public Sub SpecifyKind_WhenDifferentAsCurrentValue_ReturnsNewInstance()
    Dim dt As CorDateTime
    Set dt = NewDate(#1/1/2000#, LocalKind)
    
    Dim Actual As CorDateTime
    Set Actual = CorDateTime.SpecifyKind(dt, UtcKind)
    
    Assert.That Actual, Iz.Not.SameAs(dt)
    Assert.That Actual.Kind, Iz.EqualTo(UtcKind)
End Sub

Public Sub SpecifyKind_WithNothing_ThrowsArgumentNullException()
    On Error Resume Next
    
    CorDateTime.SpecifyKind Nothing, UtcKind
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub ToOADate_WithDefaultDate_ReturnsMidnight()
    Dim SUT As New CorDateTime
    
    Dim Actual As Date
    Actual = SUT.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#12:00:00 AM#)
End Sub

Public Sub ToOADate_WithValues_ReturnsExpected(ByVal Expected As Date)
    Dim Actual As Date
    
    Actual = NewDate(Expected).ToOADate
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToOADate_WithFractionalPortion_RoundsCorrectly()
    Const Milliseconds As Currency = 108481890561884.3912@
    Dim dt As CorDateTime
    Dim Actual As Date
    
    Set dt = CorDateTime.FromMilliseconds(Milliseconds)
    Actual = dt.ToOADate
    
    Assert.That Actual, Iz.EqualTo(#8/28/3438 10:29:21 AM#)
End Sub

Public Sub FromOADate_WithValue_ReturnsDateTime(ByVal d As Date)
    Dim Actual As CorDateTime
        
    Set Actual = CorDateTime.FromOADate(d)
            
    Assert.That Actual.Year, Iz.EqualTo(Year(d))
    Assert.That Actual.Month, Iz.EqualTo(Month(d))
    Assert.That Actual.Day, Iz.EqualTo(Day(d))
    Assert.That Actual.Hour, Iz.EqualTo(Hour(d))
    Assert.That Actual.Minute, Iz.EqualTo(Minute(d))
    Assert.That Actual.Second, Iz.EqualTo(Second(d))
    Assert.That Actual.Millisecond, Iz.EqualTo(0)
End Sub

Public Sub MinValue_WhenCalled_ReturnsJanuaryFirstYearOne()
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.MinValue
    
    Assert.That Actual, Equals(NewDateTime(1, 1, 1))
End Sub

Public Sub MaxValue_WhenCalled_ReturnsLastMomentOfYear9999()
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.MaxValue
    
    Const MaxMilliseconds As Currency = 315537897599999.9999@
    Dim Expected As CorDateTime
    Set Expected = CorDateTime.FromMilliseconds(MaxMilliseconds)

    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByVal Other As Variant, ByVal Expected As Boolean)
    Dim SUT As CorDateTime
    Set SUT = NewDateTime(2004, 1, 2, 3, 4, 5)
    
    Dim Actual As Boolean
    Actual = SUT.Equals(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithMultipleArguments_ReturnsExpected(ByVal Other As Variant, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CorDateTime.Equals(#1/2/2001#, Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub IsLeapYear_WithValue_ReturnsExpected(ByVal Year As Long, ByVal Expected As Boolean)
    Dim Actual As Boolean
    
    Actual = CorDateTime.IsLeapYear(Year)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewDateTime_WithTimeComponentsOnly_SetsToOADate()
    Dim dt As CorDateTime
    
    Set dt = NewDateTime(Hour:=18, Minute:=39, Second:=52)
    
    Assert.That dt.ToOADate, Iz.EqualTo(#6:39:52 PM#)
End Sub

Public Sub DayOfYear_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = NewDate(d).DayOfYear
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DayOfWeek_WithValue_ReturnsExpected(ByVal d As Date, ByVal Expected As DayOfWeek)
    Dim Actual As DayOfWeek
    
    Actual = NewDate(d).DayOfWeek
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TotalMilliseconds_WhenCalled_ReturnsMilliseconds()
    Dim Actual As Currency
    
    Actual = NewDateTime(2004, 1, 1, 1, 1, 1, 1).TotalMilliseconds
    
    Const Expected As Currency = 63208515661001@
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Ticks_WhenCalled_ReturnsTicks()
    Dim Actual As Variant
    
    Actual = NewDateTime(2004, 1, 1, 1, 1, 1, 1).Ticks
    
    Assert.That Actual, Iz.EqualTo(CDec("632085156610010000"))
End Sub

Public Sub DateOnly_WhenCalled_ReturnsDateTimeWithDatePartOnly()
    Dim Actual As CorDateTime
        
    Set Actual = NewDateTime(2004, 1, 1, 1, 1, 1, 1).DateOnly
    
    Assert.That Actual, Equals(#1/1/2004#)
End Sub

Public Sub TimeOfDay_WhenCalled_ReturnsExpected()
    Dim Actual As TimeSpan
    
    Set Actual = NewDateTime(2004, 1, 1, 6, 10, 15, 100).TimeOfDay
    
    Dim Expected As TimeSpan
    Set Expected = NewTimeSpan(Hours:=6, Minutes:=10, Seconds:=15, Milliseconds:=100)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Add_WithOneHour_ReturnsDateTimePlus1Hour()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    Dim ts As TimeSpan
    Set ts = TimeSpan.FromHours(1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.Add(ts)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 1, 1)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMilliseconds_With100_ReturnsDateTimePlus100Milliseconds()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddMilliseconds(100)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 1, Millisecond:=100)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddDays_With5_ReturnsDateTimePlus5Days()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddDays(5)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 6)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddHours_With5_ReturnsDateTimePlus5Hours()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddHours(5)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 1, 5)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMinutes_With5_ReturnsDateTimePlus5Minutes()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddMinutes(5)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 1, Minute:=5)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddSeconds_With60_ReturnsDateTimePlus1Minute()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2000, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddSeconds(60)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2000, 1, 1, Minute:=1)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddMonths_WithValues_ReturnsExpected(ByVal Months As Long, ByVal Expected As CorDateTime)
    Dim dt As CorDateTime
    Set dt = NewDateTime(2001, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddMonths(Months)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub AddYear_With5_ReturnsDateTimePlus5Years()
    Dim dt As CorDateTime
    Set dt = NewDateTime(2004, 5, 16)
    
    Dim Actual As CorDateTime
    Set Actual = dt.AddYears(5)
    
    Dim Expected As CorDateTime
    Set Expected = NewDateTime(2009, 5, 16)
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub DaysInMonth_WithValues_ReturnsExpected(ByVal Year As Long, ByVal Month As Long, ByVal Expected As Long)
    Dim Actual As Long
    
    Actual = CorDateTime.DaysInMonth(Year, Month)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DaysInMonth_WithInvalidValues_ThrowsArgumentOutOfRangeException(ByVal Year As Long, ByVal Month As Long, ByVal ParameterName As String)
    On Error Resume Next
    
    CorDateTime.DaysInMonth Year, Month
    
    AssertArgumentOutOfRangeException Err, ParameterName
End Sub

Public Sub CompareTo_WithValues_ReturnsExpected(ByRef Other As Variant, ByVal Expected As Long)
    Dim This As CorDateTime
    Set This = NewDate(#1/1/2001#)
    
    Dim Actual As Long
    Actual = This.CompareTo(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub CompareTo_WithNothing_Returns1()
    Dim Actual As Long
    
    Actual = CorDateTime.MinValue.CompareTo(Nothing)
    
    Assert.That Actual, Iz.EqualTo(1)
End Sub

Public Sub CompareTo_WithUnsupportedType_ThrowsArgumentException()
    On Error Resume Next
    
    NewDate(#1/1/2001#).CompareTo Empty
    
    AssertArgumentException Err
End Sub

Public Sub Subtract_WithTimeSpans_ReturnsExpected(ByVal TimeSpanToSubtract As TimeSpan, ByVal Expected As CorDateTime)
    Dim dt As CorDateTime
    Set dt = NewDateTime(2001, 1, 1)
    
    Dim Actual As CorDateTime
    Set Actual = dt.Subtract(TimeSpanToSubtract)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Subtract_WithDate_ReturnsExpected(ByVal DateToSubtract As Variant, ByVal Expected As TimeSpan)
    Dim dt As CorDateTime
    Set dt = NewDate(#1/1/2001#)
    
    Dim Actual As TimeSpan
    Set Actual = dt.Subtract(DateToSubtract)
    
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub Subtract_WithNothing_ThrowsArgumentNullException()
    Dim dt As CorDateTime
    Set dt = NewDate(#1/1/2001#)
    
    On Error Resume Next
    dt.Subtract Nothing
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub Subtract_WithNonSupportedObject_ThrowsArgumentException()
    Dim dt As CorDateTime
    Set dt = NewDate(#1/1/2001#)
    
    On Error Resume Next
    Dim Value As New Collection
    dt.Subtract Value
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub Now_WhenCalled_ReturnsLocalTime()
    Dim dt As CorDateTime

    Set dt = CorDateTime.Now
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.LocalKind)
End Sub

Public Sub Now_WhenCalled_ReturnsCurrentDateTime()
    Dim Actual As Long
    
    Actual = CorDateTime.Now.Subtract(NewDateTime(1970, 1, 1)).TotalSeconds
    
    Dim Diff As Long
    Diff = DateDiff("s", #1/1/1970#, Now) - Actual
    Assert.That Diff, Iz.GreaterThanOrEqualTo(-5).AndAlso.LessThanOrEqualTo(5)
End Sub

Public Sub UtcNow_WhenCalled_ReturnsUtcTime()
    Dim dt As CorDateTime
    
    Set dt = CorDateTime.UtcNow
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.UtcKind)
End Sub

Public Sub Today_WhenCalled_ReturnsLocalTime()
    Dim dt As CorDateTime
    
    Set dt = CorDateTime.Today
    
    Assert.That dt.Kind, Iz.EqualTo(DateTimeKind.LocalKind)
End Sub

Public Sub Today_WhenCalled_ReturnsCurrentDate()
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.Today
    
    Assert.That Actual.Equals(Date), Iz.True
End Sub

Public Sub Parse_WithDate_ReturnsParsedDateTime()
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.Parse("1/1/2004 8:30:30am")
    
    Assert.That Actual, Equals(NewDate(#1/1/2004 8:30:30 AM#))
End Sub

Public Sub FromFileTimeUtc_WhenCalled_ReturnsExpected()
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.FromFileTimeUtc(12.3456)
    
    Assert.That Actual.Kind, Iz.EqualTo(DateTimeKind.UtcKind)
    Assert.That Actual.Ticks, Iz.EqualTo(CDec("504911232000123456"))
End Sub

Public Sub FromFileTimeUtc_WithNegativeValue_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    CorDateTime.FromFileTimeUtc -1
    
    AssertArgumentOutOfRangeException Err, "Time"
End Sub

Public Sub ToFileTime_WithKinds_ReturnsExpected(ByVal Kind As DateTimeKind, ByVal Expected As Currency)
    AssertPacificTimeZone
    
    Dim Actual As Currency
    
    Actual = NewDateTime(2015, 6, 19, 10, 0, 0, Kind:=Kind).ToFileTime
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToFileTime_WhenLessThanMinAllowed_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    NewDate(#12/31/1600 11:59:59 PM#, UtcKind).ToFileTime
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToFileTimeUtc_WithKinds_ReturnsExpected(ByVal Kind As DateTimeKind, ByVal Expected As Currency)
    AssertPacificTimeZone
    
    Dim Actual As Currency
    
    Actual = NewDateTime(2015, 6, 19, 10, 0, 0, Kind:=Kind).ToFileTimeUtc
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToLocalTime_WithLocalKind_ReturnsSameValue()
    Dim d As CorDateTime
    Set d = NewDate(#1/1/2001 6:00:00 AM#, LocalKind)
    
    Dim Actual As CorDateTime
    Set Actual = d.ToLocalTime
    
    Assert.That Actual, Equals(#1/1/2001 6:00:00 AM#)
End Sub

Public Sub ToLocalTime_WithKinds_ReturnsLocalKind(ByVal k As DateTimeKind)
    Dim d As CorDateTime
    Set d = NewDate(#1/1/2001 6:00:00 AM#, k)
    
    Dim Actual As DateTimeKind
    Actual = d.ToLocalTime.Kind
    
    Assert.That Actual, Iz.EqualTo(LocalKind)
End Sub

Public Sub ToUniversalTime_WithUtcKind_ReturnsSameValue()
    Dim d As CorDateTime
    Set d = NewDate(#1/1/2001 6:00:00 AM#, UtcKind)
    
    Dim Actual As CorDateTime
    Set Actual = d.ToUniversalTime
    
    Assert.That Actual, Equals(#1/1/2001 6:00:00 AM#)
End Sub

Public Sub ToUniversalTime_WithKinds_ReturnsUtcKind(ByVal k As DateTimeKind)
    Dim d As CorDateTime
    Set d = NewDate(#1/1/2001 6:00:00 AM#, k)
    
    Dim Actual As DateTimeKind
    Actual = d.ToUniversalTime.Kind
    
    Assert.That Actual, Iz.EqualTo(UtcKind)
End Sub

Public Sub FromBinary_WhenCalled_ReturnsExpected()
    AssertPacificTimeZone
    
    Dim Actual As CorDateTime
    
    Set Actual = CorDateTime.FromBinary(-915885960650477.5808@)
    
    Assert.AreEqual DateTimeKind.LocalKind, Actual.Kind
    Assert.AreEqual 205, Actual.Year
    Assert.AreEqual 6, Actual.Month
    Assert.AreEqual 7, Actual.Day
    Assert.AreEqual 20, Actual.Hour
    Assert.AreEqual 57, Actual.Minute
    Assert.AreEqual 15, Actual.Second
End Sub

Public Sub ToBinary_WhenCalled_ReturnsCompositeTicksAndKind()
    AssertPacificTimeZone
    
    Dim Actual As Currency
    
    Actual = NewDate(#6/7/205 8:57:15 PM#, LocalKind).ToBinary
    
    Assert.That Actual, Iz.EqualTo(-915885960650477.5808@)
End Sub

Public Sub ToString_WithValue_ReturnsExpected(ByVal d As CorDateTime, ByVal Expected As String)
    Dim Actual As String
    
    Actual = d.ToString(Provider:=NewCultureInfo("en-US", False))
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCultureInfo_ReturnsExpected(ByVal Format As String, ByVal CultureName As String, ByVal Expected As String)
    Dim Culture As CultureInfo
    Set Culture = NewCultureInfo(CultureName, False)
    Dim Actual As String
    
    Actual = NewDateTime(2009, 6, 15, 13, 45, 30).ToString(Format, Culture)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFormat_ReturnsExpected(ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = NewDateTime(2001, 2, 3, 13, 5, 6, 123).ToString(Format, CultureInfo.InvariantCulture)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFractionsOfSecond_ReturnsExpected(ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = CorDateTime.FromMilliseconds(91234567890765.4321@).ToString(Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub


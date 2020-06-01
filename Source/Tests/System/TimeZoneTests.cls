VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZoneTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TimeZoneTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "IsDayLightSavings_WithValues_ReturnsExpected"
            Test.Use #2/25/2014#, False
            Test.Use #5/1/2014#, True
            Test.Use #11/25/2014#, False
            Test.Use #6/1/2014#, False
            
    End Select
End Sub

Public Sub Constructor_WhenCalled_ReturnsExpected()
    Dim Zone As TimeZone
    Set Zone = TimeZone.CurrentTimeZone
    
    Select Case UCase$(Zone.StandardName)
        Case "EASTERN STANDARD TIME"
        Case "PACIFIC STANDARD TIME"
            PacificTests Zone
        Case "CENTRAL STANDARD TIME"
        Case "MOUNTAIN STANDARD TIME"
        
        Case "US MOUNTAIN STANDARD TIME"
    
        Case Else
            Assert.Fail "Time zone '" & Zone.StandardName & "' is not supported by tests."
    End Select
End Sub

Private Sub PacificTests(ByVal Zone As TimeZone)
    Assert.That Zone.StandardName, Iz.EqualTo("Pacific Standard Time")
    Assert.That Zone.DaylightName, Iz.EqualTo("Pacific Daylight Time")
    
    Dim dl As DayLightTime
    Set dl = Zone.GetDayLightChanges(2014)
    Assert.That dl.StartTime, Equals(Cor.NewDateTime(2014, 3, 9, 2, 0, 0))
    Assert.That dl.EndTime, Equals(Cor.NewDateTime(2014, 11, 2, 2, 0, 0))
    Assert.That dl.Delta, Equals(TimeSpan.FromHours(8))
    Assert.That Zone.GetUtcOffset(#2/25/2014#), Equals(TimeSpan.FromHours(-8))
    Assert.That Zone.GetUtcOffset(#5/25/2014#), Equals(TimeSpan.FromHours(-7))
    Assert.That Zone.GetUtcOffset(#11/25/2014#), Equals(TimeSpan.FromHours(-8))
    Assert.That Zone.GetUtcOffset(#6/1/2014#), Equals(TimeSpan.FromHours(-7))
End Sub

Public Sub IsDayLightSavings_WithValues_ReturnsExpected(ByVal DateToTest As Date, ByVal Expected As Boolean)
    Dim Zone As TimeZone
    Set Zone = TimeZone.CurrentTimeZone
    
    Dim Actual As Boolean
    Actual = Zone.IsDayLightSavingTime(DateToTest)
    
    If (UCase$(Zone.StandardName) <> "US MOUNTAIN STANDARD TIME") Then 'This zone doesn't observe DST
        Assert.That Actual, Iz.EqualTo(Expected)
    Else
        Assert.Pass
    End If
End Sub

Public Sub ToLocalTime_WithLocalDateTime_ReturnsSameValue()
    Dim dt As CorDateTime
    Set dt = Cor.NewDate(#1/1/2004 12:00:00 PM#, LocalKind)
    
    Dim Actual As CorDateTime
    Set Actual = TimeZone.CurrentTimeZone.ToLocalTime(dt)
    
    Assert.That Actual, Equals(dt)
End Sub

Public Sub ToUniversalTime_WithUtcDateTime_ReturnsSameValue()
    Dim dt As CorDateTime
    Set dt = Cor.NewDate(#1/1/2004 12:00:00 PM#, UtcKind)
    
    Dim Actual As CorDateTime
    Set Actual = TimeZone.CurrentTimeZone.ToUniversalTime(dt)
    
    Assert.That Actual, Equals(dt)
End Sub


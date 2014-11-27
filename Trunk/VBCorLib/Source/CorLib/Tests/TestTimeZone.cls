VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

'
' change these tests to match your specific time zone.
'

Private Zone As TimeZone


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "IsDayLightSavings_WithValues_ReturnsExpected"
            Test.Use #2/25/2014#, False
            Test.Use #5/1/2014#, True
            Test.Use #11/25/2014#, False
            
    End Select
End Sub

Public Sub TestFixtureSetup()
'    Assert.Say "Any TimeZone failures may be because of regional settings."
'    Assert.Say "Please inspect any failures for true inaccuracy."
End Sub

Public Sub Setup()
    Set Zone = TimeZone.CurrentTimeZone
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
    
        Case Else
            Assert.Fail "Time zone '" & Zone.StandardName & "' is not supported by tests."
    End Select
End Sub

Private Sub PacificTests(ByVal Zone As TimeZone)
'    DaylightTime d1 = t1.GetDaylightChanges (2002);
'Assert.AreEqual("04/07/2002 02:00:00", d1.Start.ToString ("G"), "B03");
'Assert.AreEqual("10/27/2002 02:00:00", d1.End.ToString ("G"), "B04");
'Assert.AreEqual(36000000000L, d1.Delta.Ticks, "B05");
'DaylightTime d2 = t1.GetDaylightChanges (1996);
'Assert.AreEqual("04/07/1996 02:00:00", d2.Start.ToString ("G"), "B06");
'Assert.AreEqual("10/27/1996 02:00:00", d2.End.ToString ("G"), "B07");
'Assert.AreEqual(36000000000L, d2.Delta.Ticks, "B08");
'DateTime d3 = new DateTime (2002,2,25);
'Assert.AreEqual(false, t1.IsDaylightSavingTime (d3), "B09");
'DateTime d4 = new DateTime (2002,4,8);
'Assert.AreEqual(true, t1.IsDaylightSavingTime (d4), "B10");
'DateTime d5 = new DateTime (2002,11,4);
'Assert.AreEqual(false, t1.IsDaylightSavingTime (d5), "B11");
'Assert.AreEqual(-180000000000L, t1.GetUtcOffset (d3).Ticks, "B12");
'Assert.AreEqual(-144000000000L, t1.GetUtcOffset (d4).Ticks, "B13");
'Assert.AreEqual(-180000000000L, t1.GetUtcOffset (d5).Ticks, "B14");

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
End Sub

Public Sub IsDayLightSavings_WithValues_ReturnsExpected(ByVal DateToTest As Date, ByVal Expected As Boolean)
    Dim Zone As TimeZone
    Set Zone = TimeZone.CurrentTimeZone
    
    Dim Actual As Boolean
    Actual = Zone.IsDayLightSavingTime(DateToTest)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub TestToLocalTime()
    Dim dt As cDateTime
    
    Set dt = Zone.ToLocalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 4:00:00 AM#), "ToLocalTime should return 1/1/2004 4am."
    
    Set dt = Zone.ToLocalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 5:00:00 AM#), "ToLocalTime should return 6/1/2004 5am."
End Sub

Public Sub TestToUniversalTime()
    Dim dt As cDateTime
    
    Set dt = Zone.ToUniversalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 8:00:00 PM#), "ToUniversalTime should return 1/1/2004 8pm."
    
    Set dt = Zone.ToUniversalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 7:00:00 PM#), "ToUniversalTime should return 6/1/2004 7pm."
End Sub

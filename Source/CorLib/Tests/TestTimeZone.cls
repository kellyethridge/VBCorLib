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

'
' change these tests to match your specific time zone.
'

Private zone As TimeZone

Public Sub TestFixtureSetup()
'    Assert.Say "Any TimeZone failures may be because of regional settings."
'    Assert.Say "Please inspect any failures for true inaccuracy."
End Sub

Public Sub Setup()
    Set zone = TimeZone.CurrentTimeZone
End Sub

Public Sub TestDayLightName()
    Assert.AreEqual "Pacific Daylight Time", zone.DaylightName
End Sub

Public Sub TestStandardName()
    Assert.AreEqual "Pacific Standard Time", zone.StandardName
End Sub

Public Sub TestGetUtcOffsetOutsideOfDaylightSaving()
    Dim ts As TimeSpan
    
    Set ts = zone.GetUtcOffset(#1/1/2004#)
    Assert.AreEqual -8#, ts.TotalHours, "Should have an offset of -8 hours."
    
    Set ts = zone.GetUtcOffset(#6/6/2004#)
    Assert.AreEqual -7#, ts.TotalHours, "Should have an offset of -7 hours."
End Sub

Public Sub TestIsDayLightSavings()
    Assert.IsFalse zone.IsDayLightSavingTime(#3/11/2007 1:59:00 AM#), "3/11/2007 1:59am Should not be DLS."
    Assert.IsTrue zone.IsDayLightSavingTime(#3/11/2007 2:00:00 AM#), "3/11/2007 2am Should be DLS."
    Assert.IsTrue zone.IsDayLightSavingTime(#11/4/2007 1:59:00 AM#), "11/4/2007 1:59am should be DLS."
    Assert.IsFalse zone.IsDayLightSavingTime(#11/4/2007 2:00:00 AM#), "11/4/2007 2am should not be DLS."

End Sub

Public Sub TestDayLightChanges2007()
    Dim dlt As DayLightTime
    
    Set dlt = zone.GetDayLightChanges(2007)
    Assert.AreEqual #3/11/2007 2:00:00 AM#, dlt.StartTime.ToOADate, "DLS start time should be 3/11/2007 2am"
    Assert.AreEqual #11/4/2007 2:00:00 AM#, dlt.EndTime.ToOADate, "DLS end time should be 11/4/2007 2am"
End Sub

Public Sub TestToLocalTime()
    Dim dt As cDateTime
    
    Set dt = zone.ToLocalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 4:00:00 AM#), "ToLocalTime should return 1/1/2004 4am."
    
    Set dt = zone.ToLocalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 5:00:00 AM#), "ToLocalTime should return 6/1/2004 5am."
End Sub

Public Sub TestToUniversalTime()
    Dim dt As cDateTime
    
    Set dt = zone.ToUniversalTime(#1/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#1/1/2004 8:00:00 PM#), "ToUniversalTime should return 1/1/2004 8pm."
    
    Set dt = zone.ToUniversalTime(#6/1/2004 12:00:00 PM#)
    Assert.IsTrue dt.Equals(#6/1/2004 7:00:00 PM#), "ToUniversalTime should return 6/1/2004 7pm."
End Sub

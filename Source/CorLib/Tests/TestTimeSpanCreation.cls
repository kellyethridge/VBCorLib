VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTimeSpanCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpanCreation
'
Option Explicit

Private ts As TimeSpan

Public Sub Setup()
    Set ts = Nothing
End Sub

Public Sub TestFromMilliseconds()
    Set ts = TimeSpan.FromMilliseconds(1234567)
    Assert.AreEqual 1234567@, ts.TotalMilliseconds
End Sub

Public Sub TestFromSeconds()
    Set ts = TimeSpan.FromSeconds(12345789)
    Assert.AreEqual 12345789#, ts.TotalSeconds
End Sub

Public Sub TestFromMinutes()
    Set ts = TimeSpan.FromMinutes(123456789)
    Assert.AreEqual 123456789#, ts.TotalMinutes
End Sub

Public Sub TestFromHours()
    Set ts = TimeSpan.FromHours(123456789)
    Assert.AreEqual 123456789#, ts.TotalHours
End Sub

Public Sub TestFromDays()
    Set ts = TimeSpan.FromDays(1234567)
    Assert.AreEqual 1234567#, ts.TotalDays
End Sub

Public Sub TestFromTicks()
    Set ts = TimeSpan.FromTicks(CDec("123456789012345"))
    Assert.AreEqual CDec("123456789012345"), ts.Ticks
End Sub

Public Sub TestParseHMMSS()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "   1:06:07  "
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual 1, ts.Hours
    Assert.AreEqual 6, ts.Minutes
    Assert.AreEqual 7, ts.Seconds
End Sub

Public Sub TestParseHHMMSS()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "22:14:14"
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual 22, ts.Hours
    Assert.AreEqual 14, ts.Minutes
    Assert.AreEqual 14, ts.Seconds
End Sub

Public Sub TestNegativeHHMMSS()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "-22:14:14"
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual -22, ts.Hours
    Assert.AreEqual -14, ts.Minutes
    Assert.AreEqual -14, ts.Seconds
End Sub

Public Sub TestParseDHHMMSS()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "45.5:05:05"
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual 45, ts.Days
    Assert.AreEqual 5, ts.Hours
    Assert.AreEqual 5, ts.Minutes
    Assert.AreEqual 5, ts.Seconds
End Sub

Public Sub TestParseHHMMSSFF()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "5:06:07.1"
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual 5, ts.Hours
    Assert.AreEqual 6, ts.Minutes
    Assert.AreEqual 7, ts.Seconds
    Assert.AreEqual 100, ts.Milliseconds
End Sub

Public Sub TestParseDHHMMSSFF()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "-7.15:35:41.7777777"
    Set ts = TimeSpan.Parse(s)
    Assert.AreEqual CDec(-6609417777777#), ts.Ticks
End Sub

Public Sub TestParseFormatError()
    Dim s As String
    Dim ts As TimeSpan
    
    s = "1-5"
    On Error GoTo errTrap
    Set ts = TimeSpan.Parse(s)
    On Error GoTo 0
    Assert.Fail "An error should be thrown."

errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception type caught."
End Sub

Public Sub TestTryParseDays()
    Assert.IsTrue TimeSpan.TryParse("10", ts)
    Assert.AreEqual 10, ts.TotalDays
End Sub

Public Sub TestTryParseDaysAndHoursFails()
    Assert.IsFalse TimeSpan.TryParse("1.2", ts)
End Sub

Public Sub TestTryParseHoursAndMinutes()
    Assert.IsTrue TimeSpan.TryParse("2:3", ts)
    Assert.AreEqual 2, ts.Hours
    Assert.AreEqual 3, ts.Minutes
End Sub

Public Sub TestParseHMM()
    Dim ts As TimeSpan
    Set ts = TimeSpan.Parse("5:05")
    Assert.AreEqual 5, ts.Hours
    Assert.AreEqual 5, ts.Minutes
End Sub

Public Sub TestTryParseValidValue()
    Dim ts As TimeSpan
    Assert.IsTrue TimeSpan.TryParse("5:05", ts)
    Assert.AreEqual 5, ts.Hours
    Assert.AreEqual 5, ts.Minutes
End Sub

Public Sub TestTryParseInvalidValue()
    Dim ts As TimeSpan
    Assert.IsFalse TimeSpan.TryParse("abc", ts)
    Assert.IsNothing ts
End Sub

Public Sub TestParseOverflowValue()
    On Error GoTo errTrap
    Set ts = TimeSpan.Parse("1:60")
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is OverflowException, "Wrong exception type caught."
End Sub

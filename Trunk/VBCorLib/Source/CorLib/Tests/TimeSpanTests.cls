VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpanTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestTimeSpan
'
Option Explicit
Private ts As TimeSpan

Public Sub TestZero()
    Set ts = TimeSpan.Zero
    Assert.AreEqual 0@, ts.TotalMilliseconds, "TotalMilliseconds should return 0."
End Sub

Public Sub TestEquals()
    Dim ts1 As TimeSpan
    Dim ts2 As TimeSpan
    
    Set ts1 = TimeSpan.FromHours(2.5)
    Set ts2 = TimeSpan.FromMinutes(150)
    Assert.IsTrue ts1.Equals(ts2), "ts1 should equal ts2."
End Sub

Public Sub TestAdd()
    Dim ts1 As TimeSpan
    Dim ts2 As TimeSpan
    Dim sum As TimeSpan
    
    Set ts1 = TimeSpan.FromHours(1)
    Set ts2 = TimeSpan.FromMinutes(30)
    Set sum = ts1.Add(ts2)
    Assert.AreEqual 90#, sum.TotalMinutes, "TotalMinutes should return 90."
End Sub

Public Sub TestDuration()
    Dim negTS As TimeSpan
    Dim posts As TimeSpan
    
    Set negTS = TimeSpan.FromHours(-1.5)
    Set posts = negTS.Duration
    Assert.AreEqual 90#, posts.TotalMinutes, "TotalMinutes should return 90."
End Sub

Public Sub TestCompareTo()
    Set ts = TimeSpan.FromHours(1)
    
    Assert.AreEqual -1, ts.CompareTo(TimeSpan.FromMinutes(61)), "CompareTo should return -1."
    Assert.AreEqual 0, ts.CompareTo(TimeSpan.FromMinutes(60)), "CompareTo should return 0."
    Assert.AreEqual 1, ts.CompareTo(TimeSpan.FromMilliseconds(59)), "CompareTo should return 1."
    Assert.AreEqual 1, ts.CompareTo(Nothing), "CompareTo should return 1."
    Assert.AreEqual 1, ts.CompareTo(Null), "CompareTo should return 1."
End Sub

Public Sub TestNegate()
    Dim negTS As TimeSpan
    
    Set ts = TimeSpan.FromHours(1)
    Set negTS = ts.Negate
    Assert.AreEqual -60#, negTS.TotalMinutes, "TotalMinutes should return -60."
End Sub

Public Sub TestSubtrace()
    Dim sum As TimeSpan
    
    Set ts = TimeSpan.FromHours(1)
    Set sum = ts.Subtract(TimeSpan.FromMinutes(30))
    Assert.AreEqual 30#, sum.TotalMinutes, "TimeMinutes should return 30."
End Sub

Public Sub TestToString()
    Set ts = newtimespan(102, 3, 4, 5, 10)
    Assert.AreEqual "102.03:04:05.0100000", ts.ToString
End Sub

Public Sub TestGetHashCode()
    Dim ts2 As TimeSpan
    Dim ts3 As TimeSpan
    
    Set ts = newtimespan(1, 2, 3, 4, 5)
    Set ts2 = newtimespan(1, 2, 3, 4, 5)
    Set ts3 = newtimespan(1, 2, 3, 4, 6)
    
    Assert.IsTrue ts.GetHashCode = ts2.GetHashCode, "Hashcodes should be same."
    Assert.IsFalse ts.GetHashCode = ts3.GetHashCode, "Hashcode should be different."
End Sub

Public Sub TestNegToString()
    Set ts = newtimespan(-2, -3, -4, -5, -10)
    Assert.AreEqual "-2.03:04:05.0100000", ts.ToString
End Sub

Public Sub TestNotEqualTo()
    Dim ts1 As TimeSpan
    Dim ts2 As TimeSpan
    
    Set ts1 = TimeSpan.FromHours(1)
    Set ts2 = TimeSpan.FromDays(2)
    Assert.IsTrue ts1.NotEqualTo(ts2), "ts1 should not be equal to ts2."
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestINIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestINIFile
'
Option Explicit
Implements ICategorizable

Private ini As INIFile
Private P As String

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.IO"
End Sub

Public Sub Setup()
    P = App.Path & "\TestINI.ini"
    File.Delete P
    
    Set ini = NewINIFile(P)
End Sub

Public Sub Teardown()
    File.Delete P
End Sub

Public Sub TestSetGetString()
    ini.SetValue "test", "value", "hello"
    Assert.AreEqual "hello", ini.GetString("test", "value")
End Sub

Public Sub TestGetStringMissing()
    Assert.AreEqual "", ini.GetString("test", "missing")
End Sub

Public Sub TestGetStringWithDefault()
    Assert.AreEqual "mydefault", ini.GetString("test", "missing", "mydefault")
End Sub

Public Sub TestGetStringEmpty()
    ini.SetValue "test", "empty", ""
    Assert.AreEqual "", ini.GetString("test", "empty", "notempty")
End Sub

Public Sub TestGetLong()
    ini.SetValue "test", "long", 1
    Assert.AreEqual 1, ini.GetLong("test", "long")
End Sub

Public Sub TestGetLongMissing()
    Assert.AreEqual 0, ini.GetLong("test", "missing")
End Sub

Public Sub TestGetLongDefault()
    Assert.AreEqual 22, ini.GetLong("test", "missing", 22)
End Sub

Public Sub TestSetGetDateTime()
    Dim dt As CorDateTime
    
    Set dt = NewDate(#1/1/2005 8:30:00 AM#)
    ini.SetValue "test", "date", dt
    
    Dim dt2 As CorDateTime
    Set dt2 = ini.GetDateTime("test", "date")
    
    Assert.IsTrue dt.Equals(dt2)
End Sub

Public Sub TestSetGetTimeSpan()
    Dim ts As TimeSpan
    
    Set ts = NewTime(#9:15:42 AM#)
    ini.SetValue "test", "time", ts
    
    Dim ts2 As TimeSpan
    Set ts2 = ini.GetTimeSpan("test", "time")
    
    Assert.IsTrue ts.Equals(ts2)
End Sub

Public Sub TestGetTimeSpanDefault()
    Assert.AreEqual TimeSpan.MinValue, ini.GetTimeSpan("test", "missing", TimeSpan.MinValue)
End Sub

Public Sub TestGetSectionNames()
    Dim s() As String
    Dim i As Long
    
    For i = 1 To 10
        ini.SetValue "Section " & i, "Key " & i, "Value " & i
    Next i
    
    s = ini.GetSectionNames
    
    Assert.AreEqual 10, CorArray.GetLength(s)
    For i = 1 To 10
        Assert.AreEqual "Section " & i, s(i - 1)
    Next i
    
End Sub

Public Sub TestGetSectionNamesEmpty()
    Dim s() As String
    
    s = ini.GetSectionNames
    
    Assert.AreEqual 0, CorArray.GetLength(s)
End Sub

Public Sub TestDeleteKey()
    ini.SetValue "test", "deleteme,", "value"
    
    ini.DeleteKey "test", "deleteme"
    Assert.AreEqual "defaultvalue", ini.GetString("test", "deleteme", "defaultvalue")
End Sub

Public Sub TestDeleteSection()
    Dim s() As String
    
    ini.SetValue "deleteme", "key", "value"
    s = ini.GetSectionNames
    Assert.AreEqual 1, CorArray.GetLength(s)
    
    ini.DeleteSection "deleteme"
    s = ini.GetSectionNames
    Assert.AreEqual 0, CorArray.GetLength(s)
End Sub

Public Sub TestGetValues()
    Dim ht As Hashtable
    Dim i As Long
    
    For i = 1 To 10
        ini.SetValue "test", "key " & i, "value " & i
    Next i
    
    Set ht = ini.GetValues("test")
    
    Assert.AreEqual 10, ht.Count
    For i = 1 To 10
        Assert.IsTrue ht.Contains("key " & i)
        Assert.AreEqual "value " & i, ht("key " & i)
    Next i
End Sub

Public Sub TestSectionWriter()
    Dim w As INISectionWriter
    Dim i As Long
    Dim ht As Hashtable
    
    ini.SetValue "my section", "a key", "a value"
    
    Set w = ini.GetSectionWriter("my section")
    
    For i = 1 To 10
        w.AddResource "Key " & i, "Value " & i
    Next i
    w.Generate
    
    Set ht = ini.GetValues("my section")
    Assert.IsFalse ht.Contains("a key")
    Assert.AreEqual 10, ht.Count
    For i = 1 To 10
        Assert.IsTrue ht.Contains("key " & i)
    Next i
    
End Sub

Public Sub TestGetKeyNames()
    Dim i As Long
    Dim s() As String
    
    For i = 1 To 10
        ini.SetValue "test", "Key " & i, "Value " & i
    Next i
    
    s = ini.GetKeyNames("test")
    
    Assert.AreEqual 10, CorArray.GetLength(s)
    For i = 0 To 9
        Assert.AreEqual "Key " & i + 1, s(i)
    Next i
End Sub

Public Sub TestGetKeyNamesNone()
    Dim s() As String
    
    s = ini.GetKeyNames("missing")
    Assert.IsFalse CorArray.IsNull(s)
    Assert.AreEqual 0, CorArray.GetLength(s)
End Sub

Public Sub TestGetVeryLongValue()
    Dim s As String
    
    ini.SetValue "test", "long", String$(20000, "a")
    
    s = ini.GetString("test", "long")
    Assert.AreEqual 20000, Len(s)
    Assert.AreEqual String$(20000, "a"), s
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestHijriCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestHijriCalendar
'
Option Explicit
Implements ICategorizable

Private cal As HijriCalendar

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub Setup()
    Set cal = New HijriCalendar
End Sub

Public Sub TestIsLeapYear()
    Assert.IsFalse cal.IsLeapYear(2000)
    Assert.IsTrue cal.IsLeapYear(2001)
    Assert.IsFalse cal.IsLeapYear(2002)
End Sub

Public Sub TestGetYear()
    Dim sr As StreamReader
    Set sr = Cor.NewStreamReader(Path.Combine(App.Path, "dummy test files\HijriCalendar_GetYear.txt"))
    
    Dim s As String
    Dim d As Date
    d = #1/1/1000#
    
    s = sr.ReadLine
    Do While Len(s) > 0
        Assert.AreEqual CLng(s), cal.GetYear(d)
        d = DateAdd("d", 1, d)
        s = sr.ReadLine
    Loop
    sr.CloseReader
End Sub

Public Sub TestGetMonth()
    Dim sr As StreamReader
    Dim s As String
    Dim d As Date
    
    Set sr = Cor.NewStreamReader(Path.Combine(App.Path, "dummy test files\hijricalendar_getmonth.txt"))
    d = #1/1/1000#
    s = sr.ReadLine
    Do While Len(s) > 0
        Assert.AreEqual CLng(s), cal.GetMonth(d)
        d = DateAdd("d", 1, d)
        s = sr.ReadLine
    Loop
    
    sr.CloseReader
End Sub

Public Sub TestGetDayOfMonth()
    Assert.AreEqual 6, cal.GetDayOfMonth(#1/1/2001#)
    Assert.AreEqual 19, cal.GetDayOfMonth(#4/1/2002#)
    
    Dim sr As StreamReader
    Dim d As Date
    Dim s As String
    
    Set sr = Cor.NewStreamReader(Path.Combine(App.Path, "dummy test files\hijricalendar_getdayofmonth.txt"))
    d = #1/1/1000#
    s = sr.ReadLine
    Do While Len(s) > 0
        Assert.AreEqual CLng(s), cal.GetDayOfMonth(d)
        d = DateAdd("d", 1, d)
        s = sr.ReadLine
    Loop
    sr.CloseReader
End Sub

Public Sub TestGetDayOfYear()
    Assert.AreEqual 272, cal.GetDayOfYear(#1/1/2001#)
    Assert.AreEqual 19, cal.GetDayOfYear(#4/1/2002#)
End Sub

Public Sub TestToDateTime()
    Assert.AreEqual #12/27/2562 1:01:01 AM#, cal.ToDateTime(2001, 1, 1, 1, 1, 1, 1).ToOADate
    Assert.AreEqual #3/15/2564 1:01:01 AM#, cal.ToDateTime(2002, 4, 1, 1, 1, 1, 1).ToOADate
End Sub

Public Sub TestAddYears()
    Assert.AreEqual #12/21/2000#, cal.AddYears(#1/1/2000#, 1).ToOADate
    Assert.AreEqual #3/22/2003#, cal.AddYears(#4/1/2002#, 1).ToOADate
End Sub

Public Sub TestAddMonths()
    Assert.AreEqual #1/31/1000#, cal.AddMonths(#1/1/1000#, 1).ToOADate
    Assert.AreEqual #12/30/1000#, cal.AddMonths(#12/1/1000#, 1).ToOADate
End Sub

Public Sub TestAddDays()
    Assert.AreEqual #9/27/2002#, cal.AddDays(#1/1/2000#, 1000).ToOADate
    Assert.AreEqual #12/26/2004#, cal.AddDays(#4/1/2002#, 1000).ToOADate
End Sub

Public Sub TestDayOfYear()
    Dim d As CorDateTime
    Dim sr As StreamReader
    Dim i As Long
    
    Set sr = File.OpenText(App.Path & "\Dummy Test Files\TesthijriCalendar_DayOfYear.txt")
    Set d = CorDateTime.FromOADate(#1/1/2000#)
    
    For i = 1 To 500
        Assert.AreEqual CLng(sr.ReadLine), cal.GetDayOfYear(d)
        Set d = cal.AddDays(d, 1)
    Next i
    sr.CloseReader
End Sub

Public Sub TestGetWeekOfYear()
    Dim sr As StreamReader
    Dim i As Long
    Dim dt As CorDateTime
    Dim j As Long

    Set sr = NewStreamReader(App.Path & "\Dummy Test Files\TestHijriCalendar_WeekOfYear.txt")
    Set dt = CorDateTime.FromOADate(#1/1/2000#)

    ' tests using CalendarWeekRule.FirstDay and first day of week of DayOfWeek.Friday
    For i = 1 To 366
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Friday), "Wrong week for " & dt.ToString & " during first test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFourDayWeek and first day of week of DayOfWeek.Thursday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday), "Wrong week for " & dt.ToString & " during second test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFullWeek and first day of week of DayOfWeek.Wednesday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, cal.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday), "Wrong week for " & dt.ToString & " during third test."
        Set dt = cal.AddDays(dt, 1)
    Next i

    sr.CloseReader
End Sub

Public Sub TestLeapMonth()
    Assert.AreEqual 12, cal.GetMonth(#9/22/2082#)
End Sub

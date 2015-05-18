VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestJulianCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestJulianCalendar
'
Option Explicit
Implements ICategorizable

Private jc As JulianCalendar

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub Setup()
    Set jc = New JulianCalendar
End Sub

Public Sub TestAddDays()
    Dim jcdt As cDateTime
    Dim gcdt As cDateTime
    Dim i As Long
    
    Set jcdt = cDateTime.MinValue
    Set gcdt = cDateTime.MinValue
    
    For i = 1 To 2000
        Assert.IsTrue jcdt.EqualTo(gcdt), CorString.Format("JC {0} <> GC {1}", jcdt, gcdt)
        Set jcdt = jc.AddDays(jcdt, 1)
        Set gcdt = gcdt.AddDays(1)
    Next i
End Sub

Public Sub TestAddMonths()
    Dim dt As cDateTime
    Dim i As Long
    
    Set dt = cDateTime.FromOADate(#1/31/1000#)
    For i = 1 To 100
        Set dt = jc.AddMonths(dt, i)
    Next i
    Assert.AreEqual #12/5/1420#, dt.ToOADate
End Sub

Public Sub TestAddMonthsOnLeapYear()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/31/1000#)
    Set dt = jc.AddMonths(dt, 1)
    Assert.AreEqual 3, dt.Month, "Should return month of 3."
    Assert.AreEqual 3, dt.Day, "Should return day of 3."
    Assert.AreEqual 1000, dt.Year, "Should return year of 1000."
End Sub

Public Sub TestGetDayOfMonth()
    Assert.AreEqual 26, jc.GetDayOfMonth(#3/3/1000#)
End Sub

Public Sub TestGetMonth()
    Assert.AreEqual 2, jc.GetMonth(#3/3/1000#)
End Sub

Public Sub TestGetYear()
    Assert.AreEqual 1000, jc.GetYear(#3/3/1000#)
    Assert.AreEqual 999, jc.GetYear(#1/1/1000#)
End Sub

Public Sub TestGetHour()
    Assert.AreEqual 13, jc.GetHour(#3/3/1000 1:00:00 PM#)
End Sub

Public Sub TestGetMinute()
    Assert.AreEqual 35, jc.GetMinute(#3/3/1000 1:35:00 PM#)
End Sub

Public Sub TestGetDayOfYear()
    Dim i As Long
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromOADate(#1/1/1000#)
    For i = 360 To 369
        Assert.AreEqual (i Mod 365) + 1, jc.GetDayOfYear(dt)
        Set dt = jc.AddDays(dt, 1)
    Next i
    
    Assert.AreEqual 359, jc.GetDayOfYear(#12/31/1001#)
End Sub

Public Sub TestToDateTime()
    Dim dt As cDateTime
    
    Set dt = jc.ToDateTime(1000, 2, 29, 1, 1, 1, 1)
    Assert.AreEqual 3, dt.Month
    Assert.AreEqual 6, dt.Day
    Assert.AreEqual 1000, dt.Year
End Sub

Public Sub TestGetWeekOfYear()
    Dim sr As StreamReader
    Dim i As Long
    Dim dt As cDateTime
    Dim j As Long
    
    Set sr = NewStreamReader(App.Path & "\Dummy Test Files\TestJulianCalendar_WeekOfYear.txt")
    Set dt = cDateTime.FromOADate(#1/1/1000#)
    
    ' tests using CalendarWeekRule.FirstDay and first day of week of DayOfWeek.Friday
    For i = 1 To 366
        j = sr.ReadLine
        Assert.AreEqual j, jc.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Friday), "Wrong week for " & dt.ToString & " during first test."
        Set dt = jc.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFourDayWeek and first day of week of DayOfWeek.Thursday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, jc.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Thursday), "Wrong week for " & dt.ToString & " during second test."
        Set dt = jc.AddDays(dt, 1)
    Next i

    ' tests using CalendarWeekRule.FirstFullWeek and first day of week of DayOfWeek.Wednesday
    For i = 1 To 365
        j = sr.ReadLine
        Assert.AreEqual j, jc.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday), "Wrong week for " & dt.ToString & " during third test."
        Set dt = jc.AddDays(dt, 1)
    Next i
    
    sr.CloseReader
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCultureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestCultureInfo
'
Option Explicit

Public Sub TestInvariant()
    Dim c As CultureInfo
    
    Set c = Cor.NewCultureInfo("")
    Assert.AreEqual &H7F, c.LCID
    Assert.AreEqual "", c.Name
    Assert.AreEqual "Invariant Language (Invariant Country)", c.EnglishName
    Assert.AreEqual "Invariant Language (Invariant Country)", c.DisplayName
    Assert.IsFalse c.IsNeutralCulture, "Culture should not be neutral."
    Assert.AreEqual "Invariant Language (Invariant Country)", c.NativeName
    Assert.AreEqual "IVL", c.ThreeLetterISOLanguageName
    Assert.AreEqual "IVL", c.ThreeLetterWindowsLanguageName
    Assert.AreEqual "iv", c.TwoLetterISOLanguageName
    Assert.IsFalse c.UseUserOverride, "UseUserOverride should return false."
    
    Dim dt As DateTimeFormatInfo
    Dim i As Long
    Dim s() As String
    
    Set dt = c.DateTimeFormat
    s = dt.AbbreviatedDayNames
    Assert.AreEqual 7, cArray.GetLength(s), "AbbreviatedDayNames should return 7 elements."
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), s(i)
    Next i
    
    Assert.AreEqual "dddd, dd MMMM yyyy", dt.LongDatePattern, "Incorrect LongDatePattern."
    Assert.AreEqual "HH:mm:ss", dt.LongTimePattern, "Incorrect LongTimePattern."
    
    Dim nm As NumberFormatInfo
    Dim n() As Long
    
    Set nm = c.NumberFormat
    Assert.AreEqual ChrW$(&H2030), nm.PerMilleSymbol
    
    n = nm.NumberGroupSizes
    Assert.AreEqual 1, cArray.GetLength(n)
    Assert.AreEqual 3, n(0)
End Sub

Public Sub TestEnglishUS()
    Dim c As CultureInfo
    
    Set c = Cor.NewCultureInfo("en-US")
    Assert.AreEqual 1033, c.LCID
    Assert.AreEqual "en-US", c.Name
    Assert.AreEqual "English (United States)", c.EnglishName
    Assert.AreEqual "English (United States)", c.DisplayName
    Assert.IsFalse c.IsNeutralCulture, "Culture should not be neutral."
    Assert.AreEqual "English (United States)", c.NativeName
    Assert.AreEqual "eng", c.ThreeLetterISOLanguageName
    Assert.AreEqual "ENU", c.ThreeLetterWindowsLanguageName
    Assert.AreEqual "en", c.TwoLetterISOLanguageName
    
    Dim dt As DateTimeFormatInfo
    Dim i As Long
    Dim s() As String
    
    Set dt = c.DateTimeFormat
    s = dt.DayNames
    Assert.AreEqual 7, cArray.GetLength(s), "AbbreviatedDayNames should return 7 elements."
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"), s(i)
    Next i
    
    Assert.AreEqual "dddd, MMMM dd, yyyy", dt.LongDatePattern, "Incorrect LongDatePattern."
    Assert.AreEqual "h:mm:ss tt", dt.LongTimePattern, "Incorrect LongTimePattern."
    
End Sub

Public Sub TestDefaultCalendar()
    Dim cal As Calendar
    Dim cul As CultureInfo
    
    Set cul = Cor.NewCultureInfo("ar-AE")
    Assert.IsNotNothing cul
    
    Set cal = cul.Calendar
    Assert.IsTrue TypeOf cal Is GregorianCalendar, "Should return a gregorian calendar."
End Sub

Public Sub TestOptionalCalendars()
    Dim cals() As Calendar
    Dim cul As CultureInfo
    
    Set cul = Cor.NewCultureInfo("ar-AE")
    cals = cul.OptionalCalendars
    
    Assert.AreEqual 6, cArray.GetLength(cals)
    
    Assert.IsTrue ContainsCalendar(cals, 1)
    Assert.IsTrue ContainsCalendar(cals, 2)
    Assert.IsTrue ContainsCalendar(cals, 6)
    Assert.IsTrue ContainsCalendar(cals, 9)
    Assert.IsTrue ContainsCalendar(cals, 12)
    Assert.IsTrue ContainsCalendar(cals, 10)
    
End Sub



Private Function ContainsCalendar(ByRef Calendars() As Calendar, ByVal ID As Long) As Boolean
    Dim i As Long
    Dim g As GregorianCalendar
    
    For i = LBound(Calendars) To UBound(Calendars)
        Select Case LCase$(TypeName(Calendars(i)))
            Case "gregoriancalendar"
                Set g = Calendars(i)
                If g.CalendarType = ID Then
                    ContainsCalendar = True
                    Exit Function
                End If
            Case "hijricalendar"
                If ID = 6 Then
                    ContainsCalendar = True
                    Exit Function
                End If
        End Select
    Next i
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestDateTimeFormatInfoInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TestDateTimeFormatInfoInvariant
'
Option Explicit
Implements ICategorizable

Private info As DateTimeFormatInfo
Private i As Long
Private s() As String

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Public Sub Setup()
    Set info = New DateTimeFormatInfo
End Sub

Public Sub TestAbbreviatedDayNames()
    s = info.AbbreviatedDayNames
    Assert.AreEqual 7, CorArray.GetLength(s), "Should return 7 days."
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), s(i)
    Next i
End Sub

Public Sub TestAbbreviatedMonthsNames()
    s = info.AbbreviatedMonthNames
    Assert.AreEqual 13, CorArray.GetLength(s), "Should return 13 months."
    For i = 0 To 12
        Assert.AreEqual Choose(i + 1, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""), s(i)
    Next i
End Sub

Public Sub TestAMDesignator()
    Assert.AreEqual "AM", info.AMDesignator
End Sub

Public Sub TestDateSeparator()
    Assert.AreEqual "/", info.DateSeparator
End Sub

Public Sub TestDayNames()
    s = info.DayNames
    Assert.AreEqual 7, CorArray.GetLength(s), "Should return 7 days."
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"), s(i)
    Next i
End Sub

Public Sub TestFirstDayOfWeek()
    Assert.AreEqual DayOfWeek.Sunday, info.FirstDayOfWeek
End Sub

Public Sub TestLongDatePattern()
    Assert.AreEqual "dddd, dd MMMM yyyy", info.LongDatePattern
End Sub

Public Sub TestLongTimePattern()
    Assert.AreEqual "HH:mm:ss", info.LongTimePattern
End Sub

Public Sub TestFullDateTimePattern()
    Assert.AreEqual "dddd, dd MMMM yyyy HH:mm:ss", info.FullDateTimePattern
End Sub

Public Sub TestGetAbbreviatedDayName()
    For i = 0 To 6
        Assert.AreEqual Choose(i + 1, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"), info.GetAbbreviatedDayName(i)
    Next i
End Sub

Public Sub TestGetAllDateTimePatterns_LowerD()
    s = info.GetAllDateTimePatterns("d")
    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
    Assert.AreEqual "MM/dd/yyyy", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_UpperD()
    s = info.GetAllDateTimePatterns("D")
    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
    Assert.AreEqual "dddd, dd MMMM yyyy", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_UpperF()
    s = info.GetAllDateTimePatterns("F")
    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
    Assert.AreEqual "dddd, dd MMMM yyyy HH:mm:ss", s(0)
End Sub

Public Sub TestGetAllDateTimePatterns_LowerF()
    s = info.GetAllDateTimePatterns("f")
    Assert.AreEqual 4, CorArray.GetLength(s), "Should return 4 element array."
    Assert.That s, Iz.EquivalentTo(Array("dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt"))
    
'    Dim i As Long
'    For i = 0 To UBound(s)
'        Assert.That s(i), Iz.SubsetOf(Array("dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt"))
'    Next i
End Sub

Public Sub TestGetAllPatterns_LowerG()
    Dim s() As String
    Dim i As Long
    
    s = info.GetAllDateTimePatterns("g")
    Assert.AreEqual 4, CorArray.GetLength(s), "s should contain 4 elements."
    Assert.That s, Iz.EquivalentTo(Array("MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt"))
    
'    For i = 0 To 3
'        Assert.That s(i), Iz.SubsetOf(Array("MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt"))
'    Next i
End Sub

Public Sub TestGetAllPatterns_UpperG()
    Dim s() As String
    
    s = info.GetAllDateTimePatterns("G")
    Assert.AreEqual 1, CorArray.GetLength(s), "s should contain 1 element."
    Assert.AreEqual "MM/dd/yyyy HH:mm:ss", s(0)
End Sub

Public Sub TestGetAllPatterns()
    Dim s() As String
    Dim i As Long
    
    s = info.GetAllDateTimePatterns
    Assert.AreEqual 26, CorArray.GetLength(s), "s should contain 26 elements."
End Sub

Public Sub TestFormatDayPart()
    Assert.AreEqual "Thursday Thu 01 1", info.Format(#1/1/2004#, "dddd ddd dd d")
End Sub

Public Sub TestFormatTime()
    Assert.AreEqual "05:30:15", info.Format(#5:30:15 AM#, "HH:mm:ss")
End Sub

Public Sub TestCustomFormat()
    Assert.AreEqual "Thursday, 01 January 02004 18:30", info.Format(#1/1/2004 6:30:00 PM#, "dddd, dd MMMM yyyyy HH:mm")
End Sub

Public Sub TestIsReadOnly()
    Assert.IsFalse info.IsReadOnly, "Locally created info should not be ReadOnly."
    Assert.IsTrue CultureInfo.CurrentCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
    Assert.IsTrue CultureInfo.InvariantCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
End Sub

Public Sub TestTimeZoneFormats()
    Dim s As String
    
    s = info.Format(Time, "%z")
    Assert.AreEqual "-8", s
    
    s = info.Format(Time, "%zz")
    Assert.AreEqual "-08", s
    
    s = info.Format(Time, "%zzz")
    Assert.AreEqual "-08:00", s
    
    s = info.Format(Time, "%zzzzzzz")
    Assert.AreEqual "-08:00", s
End Sub

Public Sub TestEscapeSequence()
    Assert.AreEqual "d", info.Format(Now, "\d")
    Assert.AreEqual "\", info.Format(Now, "\\")
End Sub

Public Sub TestEscapedQuote()
    Assert.AreEqual "aa'", info.Format(Now, "'aa\''")
End Sub

Public Sub TestFractionsOfSecond()
    Dim dt As cDateTime
    
    Set dt = cDateTime.FromMilliseconds(91234567890765.4321@)
    
    Assert.AreEqual "7654321", info.Format(dt, "fffffff")
    Assert.AreEqual "765432", info.Format(dt, "ffffff")
    Assert.AreEqual "76543", info.Format(dt, "fffff")
    Assert.AreEqual "7654", info.Format(dt, "ffff")
    Assert.AreEqual "765", info.Format(dt, "fff")
    Assert.AreEqual "76", info.Format(dt, "ff")
    Assert.AreEqual "7", info.Format(dt, "%f")
End Sub

Public Sub TestPersistence()
    info.DateSeparator = "^"
    
    Dim Days() As String
    Days = Cor.NewStrings("M", "T", "W", "Th", "F", "S", "S")
    
    info.AbbreviatedDayNames = Days
    
    Dim Bag As New PropertyBag
    Bag.WriteProperty "Info", info
    
    Dim bag2 As New PropertyBag
    bag2.Contents = Bag.Contents
    
    Dim NewInfo As DateTimeFormatInfo
    Set NewInfo = bag2.ReadProperty("Info")
    
    Assert.AreEqual "^", NewInfo.DateSeparator
    Assert.AreEqual Days, NewInfo.AbbreviatedDayNames
End Sub


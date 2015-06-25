VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTimeFormatInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Testing DateTimeFormatInfo is just for ensure the correctness of the class,
' not testing that the culture information itself is correct.
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

'Private info As DateTimeFormatInfo
'Private i As Long
'Private s() As String

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetAbbreviatedDayName_WithDay_ReturnsName"
            Test.Use DayOfWeek.Sunday, "Sun"
            Test.Use DayOfWeek.Monday, "Mon"
            Test.Use DayOfWeek.Tuesday, "Tue"
            Test.Use DayOfWeek.Wednesday, "Wed"
            Test.Use DayOfWeek.Thursday, "Thu"
            Test.Use DayOfWeek.Friday, "Fri"
            Test.Use DayOfWeek.Saturday, "Sat"
            
    End Select
End Sub

Public Sub AbbreviatedDayNames_WhenCalled_ReturnsArrayOfAbbreviatedDayNames()
    Dim Actual() As String
    
    Actual = MakeInfo.AbbreviatedDayNames
    
    Assert.That Actual, Iz.EqualTo(NewStrings("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"))
End Sub

Public Sub AbbreviatedMonthNames_WhenCalled_ReturnsArrayOfAbbreviatedMonthNames()
    Dim Actual() As String
    
    Actual = MakeInfo.AbbreviatedMonthNames
    
    Assert.That Actual, Iz.EqualTo(NewStrings("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", ""))
End Sub

Public Sub AMDesignator_WhenCalled_ReturnsAM()
    Dim Actual As String
    
    Actual = MakeInfo.AMDesignator
    
    Assert.That Actual, Iz.EqualTo("AM")
End Sub

Public Sub PMDesignator_WhenCalled_ReturnsPM()
    Dim Actual As String
    
    Actual = MakeInfo.PMDesignator
    
    Assert.That Actual, Iz.EqualTo("PM")
End Sub

Public Sub DateSeparator_WhenCalled_ReturnsForwardSlash()
    Dim Actual As String
    
    Actual = MakeInfo.DateSeparator
    
    Assert.That Actual, Iz.EqualTo("/")
End Sub

Public Sub DayNames_WhenCalled_ReturnsArrayOfDayNames()
    Dim Actual() As String
    
    Actual = MakeInfo.DayNames
    
    Assert.That Actual, Iz.EqualTo(NewStrings("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"))
End Sub

Public Sub FirstDayOfWeek_WhenCalled_ReturnsSunday()
    Dim Actual As DayOfWeek
    
    Actual = MakeInfo.FirstDayOfWeek
    
    Assert.That Actual, Iz.EqualTo(DayOfWeek.Sunday)
End Sub

Public Sub LongDatePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.LongDatePattern
    
    Assert.That Actual, Iz.EqualTo("dddd, MMMM dd, yyyy")
End Sub

Public Sub LongTimePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.LongTimePattern
    
    Assert.That Actual, Iz.EqualTo("h:mm:ss tt")
End Sub

Public Sub FullDateTimePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.FullDateTimePattern
    
    Assert.That Actual, Iz.EqualTo("dddd, MMMM dd, yyyy h:mm:ss tt")
End Sub

Public Sub GetAbbreviatedDayName_WithDay_ReturnsName(ByVal Day As DayOfWeek, ByVal Expected As String)
    Dim Actual As String
    
    Actual = MakeInfo.GetAbbreviatedDayName(Day)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub GetAbbreviatedDayName_WithInvalidDayOfWeek_ThrowsArgumentOutOfRangeException(ByVal Invalid As Long)
    On Error Resume Next
    
    MakeInfo.GetAbbreviatedDayName Invalid
    
    AssertArgumentOutOfRangeException Err, "DayOfWeek"
End Sub

'Public Sub TestGetAllDateTimePatterns_LowerD()
'    s = info.GetAllDateTimePatterns("d")
'    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
'    Assert.AreEqual "MM/dd/yyyy", s(0)
'End Sub
'
'Public Sub TestGetAllDateTimePatterns_UpperD()
'    s = info.GetAllDateTimePatterns("D")
'    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
'    Assert.AreEqual "dddd, dd MMMM yyyy", s(0)
'End Sub
'
'Public Sub TestGetAllDateTimePatterns_UpperF()
'    s = info.GetAllDateTimePatterns("F")
'    Assert.AreEqual 1, CorArray.GetLength(s), "Should return 1 element array."
'    Assert.AreEqual "dddd, dd MMMM yyyy HH:mm:ss", s(0)
'End Sub
'
'Public Sub TestGetAllDateTimePatterns_LowerF()
'    s = info.GetAllDateTimePatterns("f")
'    Assert.AreEqual 4, CorArray.GetLength(s), "Should return 4 element array."
'    Assert.That s, Iz.EquivalentTo(Array("dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt"))
'
''    Dim i As Long
''    For i = 0 To UBound(s)
''        Assert.That s(i), Iz.SubsetOf(Array("dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt"))
''    Next i
'End Sub
'
'Public Sub TestGetAllPatterns_LowerG()
'    Dim s() As String
'    Dim i As Long
'
'    s = info.GetAllDateTimePatterns("g")
'    Assert.AreEqual 4, CorArray.GetLength(s), "s should contain 4 elements."
'    Assert.That s, Iz.EquivalentTo(Array("MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt"))
'
''    For i = 0 To 3
''        Assert.That s(i), Iz.SubsetOf(Array("MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt"))
''    Next i
'End Sub
'
'Public Sub TestGetAllPatterns_UpperG()
'    Dim s() As String
'
'    s = info.GetAllDateTimePatterns("G")
'    Assert.AreEqual 1, CorArray.GetLength(s), "s should contain 1 element."
'    Assert.AreEqual "MM/dd/yyyy HH:mm:ss", s(0)
'End Sub
'
'Public Sub TestGetAllPatterns()
'    Dim s() As String
'    Dim i As Long
'
'    s = info.GetAllDateTimePatterns
'    Assert.AreEqual 26, CorArray.GetLength(s), "s should contain 26 elements."
'End Sub
'
'Public Sub TestIsReadOnly()
'    Assert.IsFalse info.IsReadOnly, "Locally created info should not be ReadOnly."
'    Assert.IsTrue CultureInfo.CurrentCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
'    Assert.IsTrue CultureInfo.InvariantCulture.DateTimeFormat.IsReadOnly, "CurrentCulture supplied should be ReadOnly."
'End Sub
'
'Public Sub TestPersistence()
'    info.DateSeparator = "^"
'
'    Dim Days() As String
'    Days = Cor.NewStrings("M", "T", "W", "Th", "F", "S", "S")
'
'    info.AbbreviatedDayNames = Days
'
'    Dim Bag As New PropertyBag
'    Bag.WriteProperty "Info", info
'
'    Dim bag2 As New PropertyBag
'    bag2.Contents = Bag.Contents
'
'    Dim NewInfo As DateTimeFormatInfo
'    Set NewInfo = bag2.ReadProperty("Info")
'
'    Assert.AreEqual "^", NewInfo.DateSeparator
'    Assert.AreEqual Days, NewInfo.AbbreviatedDayNames
'End Sub

Private Function MakeInfo() As DateTimeFormatInfo
    Set MakeInfo = NewCultureInfo("en-US", False).DateTimeFormat
End Function



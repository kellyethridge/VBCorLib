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
' DateTimeFormatInfoTests
'
' Testing DateTimeFormatInfo is to ensure the correctness of the class,
' not testing that the culture information itself is correct.
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

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
            
        Case "GetAbbreviatedDayName_WithInvalidDayOfWeek_ThrowsArgumentOutOfRangeException"
            Test.Use -1
            Test.Use 7
            
        Case "GetAllDateTimePatterns_WithValue_ReturnsArrayOfPatterns"
            Test.Use "d", NewStrings("MM/dd/yyyy")
            Test.Use "D", NewStrings("dddd, dd MMMM yyyy")
            Test.Use "F", NewStrings("dddd, dd MMMM yyyy HH:mm:ss")
            Test.Use "f", NewStrings("dddd, dd MMMM yyyy HH:mm", "dddd, dd MMMM yyyy hh:mm tt", "dddd, dd MMMM yyyy H:mm", "dddd, dd MMMM yyyy h:mm tt")
            Test.Use "g", NewStrings("MM/dd/yyyy HH:mm", "MM/dd/yyyy hh:mm tt", "MM/dd/yyyy H:mm", "MM/dd/yyyy h:mm tt")
            Test.Use "G", NewStrings("MM/dd/yyyy HH:mm:ss")
            
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

Public Sub DayNames_WithInvalidArrayLength_ThrowsArgumentException()
    Dim info As New DateTimeFormatInfo
    
    On Error Resume Next
    info.DayNames = NewStrings("1", "2", "3", "4", "5", "6", "7", "8")
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub AbbreviatedDayName_WithInvalidArraylength_ThrowsArgumentException()
    Dim info As New DateTimeFormatInfo
    
    On Error Resume Next
    info.AbbreviatedDayNames = NewStrings("1", "2", "3", "4", "5", "6", "7", "8")
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub MonthNames_WithInvalidArrayLength_ThrowsArgumentException()
    Dim info As New DateTimeFormatInfo
    
    On Error Resume Next
    info.MonthNames = NewStrings("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub AbbreviatedMonthNames_WithInvalidArrayLength_ThrowsArgumentException()
    Dim info As New DateTimeFormatInfo
    
    On Error Resume Next
    info.AbbreviatedMonthNames = NewStrings("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
    
    AssertArgumentException Err, "Value"
End Sub

Public Sub FirstDayOfWeek_WhenCalled_ReturnsSunday()
    Dim Actual As DayOfWeek
    
    Actual = MakeInfo.FirstDayOfWeek
    
    Assert.That Actual, Iz.EqualTo(DayOfWeek.Sunday)
End Sub

Public Sub LongDatePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.LongDatePattern
    
    Assert.That Actual, Iz.EqualTo("dddd, dd MMMM yyyy")
End Sub

Public Sub LongTimePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.LongTimePattern
    
    Assert.That Actual, Iz.EqualTo("HH:mm:ss")
End Sub

Public Sub FullDateTimePattern_WhenCalled_ReturnsDateFormatPattern()
    Dim Actual As String
    
    Actual = MakeInfo.FullDateTimePattern
    
    Assert.That Actual, Iz.EqualTo("dddd, dd MMMM yyyy HH:mm:ss")
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

Public Sub GetAllDateTimePatterns_WithValue_ReturnsArrayOfPatterns(ByVal Format As String, ByRef Expected As Variant)
    Dim Actual() As String
    
    Actual = MakeInfo.GetAllDateTimePatterns(Format)
    
    Assert.That Actual, Iz.EquivalentTo(Expected)
End Sub

Public Sub GetAllDateTimePatterns_WhenCalled_Returns26ElementArray()
    Dim Actual() As String
    
    Actual = MakeInfo.GetAllDateTimePatterns
    
    Assert.That CorArray.Length(Actual), Iz.EqualTo(26)
End Sub

Public Sub ReadOnly_WithReadOnlyInfo_ReturnsSameInstance()
    Dim info As DateTimeFormatInfo
    Set info = MakeInfo
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.ReadOnly(info)
    
    Assert.That Actual, Iz.SameAs(info)
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub ReadOnly_WithNonReadOnlyInfo_ReturnsReadOnlyCopy()
    Dim info As New DateTimeFormatInfo
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.ReadOnly(info)
    
    Assert.That Actual, Iz.Not.SameAs(info)
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim info As New DateTimeFormatInfo
    info.DateSeparator = "^"
    info.DayNames = NewStrings("M", "T", "W", "Th", "F", "S", "S")
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = Persist(info)
    
    Assert.That Actual.DateSeparator, Iz.EqualTo("^")
    Assert.That Actual.DayNames, Iz.EqualTo(NewStrings("M", "T", "W", "Th", "F", "S", "S"))
End Sub

Public Sub GetInstance_WithFormatProvider_ReturnsDateTimeFormatInfo()
    Dim Provider As New FakeFormatProvider
    Set Provider.GetFormatResult = New DateTimeFormatInfo
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.GetInstance(Provider)
    
    Assert.That Provider.FormatType, Iz.EqualTo("DateTimeFormatInfo")
    Assert.That Actual, Iz.SameAs(Provider.GetFormatResult)
End Sub

Public Sub GetInstance_WithDateTimeFormatInfo_ReturnsSameInstance()
    Dim Provider As New DateTimeFormatInfo
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.GetInstance(Provider)
    
    Assert.That Actual, Iz.SameAs(Provider)
End Sub

Public Sub GetInstance_WithCultureInfo_ReturnsDateTimeFormatInfoFromCulture()
    Dim Culture As CultureInfo
    Set Culture = Cor.NewCultureInfo("fr-FR")
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.GetInstance(Culture)
    
    Assert.That Actual, Iz.SameAs(Culture.DateTimeFormat)
End Sub

Public Sub GetInstance_WithNoInfoAvailable_ReturnsCurrentInfo()
    Dim Expected As DateTimeFormatInfo
    Set Expected = DateTimeFormatInfo.CurrentInfo
    
    Dim Actual As DateTimeFormatInfo
    Set Actual = DateTimeFormatInfo.GetInstance(New FakeFormatProvider)
    
    Assert.That Actual, Iz.SameAs(Expected)
End Sub


Private Function MakeInfo() As DateTimeFormatInfo
    Set MakeInfo = CultureInfo.InvariantCulture.DateTimeFormat
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfoTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CultureInfoTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Globalization"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Constructor_WithUseUserOverride_ReturnsExpected"
            Test.Use True
            Test.Use False
            
        Case "Contstructor_WithEmptyString_ReturnsInvariant"
            Test.Use ""
            Test.Use CultureInfo.InvariantCulture.LCID
            
        Case "IsNeutralCulture_WithValues_ReturnsExpected"
            Test.Use "en", True
            Test.Use "en-US", False
            
    End Select
End Sub

Public Sub Constructor_WithIdentifier2067_ReturnsCultureInfo()
    Dim info As CultureInfo
    
    Set info = Cor.NewCultureInfo(2067)
    
    Assert.That info.IsReadOnly, Iz.False, "Should be read-only."
    Assert.That info.LCID, Iz.EqualTo(2067)
    Assert.That info.Name, Iz.EqualTo("nl-BE")
    Assert.That info.UseUserOverride, Iz.True, "Should be user override."
End Sub

Public Sub Constructor_WithNegativeIdentifier_ThrowsArgumentOutOfRangeException()
    On Error Resume Next
    
    Cor.NewCultureInfo -1
    
    AssertArgumentOutOfRangeException Err, "LCID"
End Sub

Public Sub Constructor_WithCultureName_ReturnsCultureInfo()
    Dim info As CultureInfo
    
    Set info = Cor.NewCultureInfo("nl-BE")
    
    Assert.That info.IsReadOnly, Iz.False
    Assert.That info.LCID, Iz.EqualTo(2067)
    Assert.That info.Name, Iz.EqualTo("nl-BE")
    Assert.That info.UseUserOverride, Iz.True
End Sub

Public Sub Constructor_WithUseUserOverride_ReturnsExpected(ByVal UseUserOverride As Boolean)
    Dim info As CultureInfo
    
    Set info = Cor.NewCultureInfo(2067, UseUserOverride)
    
    Assert.That info.IsReadOnly, Iz.False
    Assert.That info.LCID, Iz.EqualTo(2067)
    Assert.That info.Name, Iz.EqualTo("nl-BE")
    Assert.That info.UseUserOverride, Iz.EqualTo(UseUserOverride)
End Sub

Public Sub DateTimeFormat_WithNeutralCulture_ThrowsNotSupportedExeption()
    Dim NeutralCulture As CultureInfo
    Set NeutralCulture = Cor.NewCultureInfo("nl")
    
    On Error Resume Next
    Dim FormatInfo As DateTimeFormatInfo
    Set FormatInfo = NeutralCulture.DateTimeFormat
    
    AssertNotSupportedException Err
End Sub

Public Sub GetAllInvariantCultures_WhenCalled_ContainsInvariantCulture()
    Dim Cultures() As CultureInfo
    Cultures = CultureInfo.GetCultures(CultureTypes.AllCultures)
    
    Dim Invariant As CultureInfo
    Set Invariant = CultureInfo.InvariantCulture
    
    Assert.That Cultures, Has.Some.EqualTo(Invariant).Using(New CorObjectComparer)
End Sub

Public Sub OptionalCalendars_WhenCalled_DoesNotReturnNull()
    Dim Item As Variant
    For Each Item In CultureInfo.GetCultures(AllCultures)
        Dim Culture As CultureInfo
        Set Culture = Item
        
        Dim Actual() As Calendar
        Actual = Culture.OptionalCalendars
        Assert.That CorArray.IsNull(Actual), Iz.False, Culture.LCID & " " & Culture.Name
    Next
End Sub

Public Sub Clone_WithNeutralCulture_ReturnsCopy()
    Dim Neutral As CultureInfo
    Set Neutral = Cor.NewCultureInfo("en")
    Dim NeutralClone As CultureInfo
    
    Set NeutralClone = Neutral.Clone
    
    Assert.That Neutral, Equals(NeutralClone)
End Sub

Public Sub Contstructor_WithEmptyString_ReturnsInvariant(ByVal NameOrLCID As Variant)
    Dim c As CultureInfo

    Set c = Cor.NewCultureInfo(NameOrLCID)
    
    Assert.That c.LCID, Iz.EqualTo(&H7F)
    Assert.That c.Name, Iz.Empty
    Assert.That c.EnglishName, Iz.EqualTo("Invariant Language (Invariant Country)")
    Assert.That c.DisplayName, Iz.EqualTo("Invariant Language (Invariant Country)")
    Assert.That c.IsNeutralCulture, Iz.False, "Culture should not be neutral."
    Assert.That c.NativeName, Iz.EqualTo("Invariant Language (Invariant Country)")
    Assert.That c.ThreeLetterISOLanguageName, Iz.EqualTo("IVL")
    Assert.That c.ThreeLetterWindowsLanguageName, Iz.EqualTo("IVL")
    Assert.That c.TwoLetterISOLanguageName, Iz.EqualTo("iv")
    Assert.That c.UseUserOverride, Iz.False, "UseUserOverride should return false."
    Assert.That c.DateTimeFormat.AbbreviatedDayNames, Iz.EqualTo(Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"))
    Assert.That c.DateTimeFormat.LongDatePattern, Iz.EqualTo("dddd, dd MMMM yyyy")
    Assert.That c.DateTimeFormat.LongTimePattern, Iz.EqualTo("HH:mm:ss")
    Assert.That c.NumberFormat.PerMilleSymbol, Iz.EqualTo(ChrW$(&H2030))
    Assert.That c.NumberFormat.NumberGroupSizes, Iz.EqualTo(Array(3))
End Sub

Public Sub Constructor_WithenUS_ReturnsUsEnglish()
    Dim c As CultureInfo

    Set c = Cor.NewCultureInfo("en-US")
    
    Assert.That c.LCID, Iz.EqualTo(1033)
    Assert.That c.Name, Iz.EqualTo("en-US")
    Assert.That c.EnglishName, Iz.EqualTo("English (United States)")
    Assert.That c.DisplayName, Iz.EqualTo("English (United States)")
    Assert.That c.IsNeutralCulture, Iz.False, "Culture should not be neutral."
    Assert.That c.NativeName, Iz.EqualTo("English (United States)")
    Assert.That c.ThreeLetterISOLanguageName, Iz.EqualTo("eng")
    Assert.That c.ThreeLetterWindowsLanguageName, Iz.EqualTo("ENU")
    Assert.That c.TwoLetterISOLanguageName, Iz.EqualTo("en")
    Assert.That c.DateTimeFormat.DayNames, Iz.EqualTo(Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"))
    Assert.That c.DateTimeFormat.LongDatePattern, Iz.EqualTo("dddd, MMMM dd, yyyy")
    Assert.That c.DateTimeFormat.LongTimePattern, Iz.EqualTo("h:mm:ss tt")
End Sub

Public Sub Calendary_WithenUsCulture_ReturnsGegorian()
    Dim c As CultureInfo
    Set c = Cor.NewCultureInfo("en-US")
    
    Dim Actual As Calendar
    Set Actual = c.Calendar
    
    Assert.That TypeOf Actual Is GregorianCalendar, Iz.True
End Sub

Public Sub IsReadOnly_WhenCreatingNewInstance_ReturnsFalse()
    Dim c As CultureInfo
    
    Set c = Cor.NewCultureInfo("en-US")
    
    Assert.That c.IsReadOnly, Iz.False
    Assert.That c.DateTimeFormat.IsReadOnly, Iz.False
    Assert.That c.NumberFormat.IsReadOnly, Iz.False
    c.NumberFormat.NumberGroupSeparator = ","
    Set c.NumberFormat = New NumberFormatInfo
    c.DateTimeFormat.DateSeparator = "/"
    Set c.DateTimeFormat = New DateTimeFormatInfo
    Assert.That c.DateTimeFormat.IsReadOnly, Iz.False
    Assert.That c.NumberFormat.IsReadOnly, Iz.False
End Sub

Public Sub IsReadOnly_WithCurrentCulture_ReturnsTrue()
    Dim c As CultureInfo
    
    Set c = CultureInfo.CurrentCulture
    
    Assert.That c.IsReadOnly, Iz.True
    Assert.That c.DateTimeFormat.IsReadOnly, Iz.True
    Assert.That c.NumberFormat.IsReadOnly, Iz.True
End Sub

Public Sub IsNeutralCulture_WithValues_ReturnsExpected(ByVal NameOrLCID As Variant, ByVal Expected As Boolean)
    Dim ci As CultureInfo
    
    Set ci = Cor.NewCultureInfo(NameOrLCID)
    
    Assert.That ci.IsNeutralCulture, Iz.EqualTo(Expected)
End Sub

Public Sub GetFormat_WithDateTimeFormatInfo_ReturnsCorrectType()
    Dim Actual As Object
    
    Set Actual = CultureInfo.CurrentCulture.GetFormat("DateTimeFormatInfo")
    
    Assert.That TypeOf Actual Is DateTimeFormatInfo, Iz.True
End Sub

Public Sub GetFormat_WithNumberFormatInfo_ReturnsCorrectType()
    Dim Actual As Object
    
    Set Actual = CultureInfo.CurrentCulture.GetFormat("NumberFormatInfo")
    
    Assert.That TypeOf Actual Is NumberFormatInfo, Iz.True
End Sub

Public Sub GetFormat_WithUnsupportedType_ReturnsNothing()
    Dim Actual As Object
    
    Set Actual = CultureInfo.CurrentCulture.GetFormat("abc")
    
    Assert.That Actual, Iz.Nothing
End Sub

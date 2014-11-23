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
Option Explicit
Implements ITestCaseSource

Private mOldCulture As CultureInfo

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "Constructor_WithUseUserOverride_ReturnsExpected"
            Test.Use True
            Test.Use False
            
        Case "Contstructor_WithEmptyString_ReturnsInvariant"
            Test.Use ""
            Test.Use CultureInfo.InvariantCulture.LCID
            
    End Select
End Sub

Public Sub Setup()
    Set mOldCulture = Thread.CurrentThread.CurrentCulture
End Sub

Public Sub TearDown()
    Set Thread.CurrentThread.CurrentCulture = mOldCulture
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
    On Error GoTo errTrap
    
    Cor.NewCultureInfo -1
    
errTrap:
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
    On Error GoTo errTrap
    Dim NeutralCulture As CultureInfo
    Set NeutralCulture = Cor.NewCultureInfo("nl")
    
    Dim FormatInfo As DateTimeFormatInfo
    Set FormatInfo = NeutralCulture.DateTimeFormat
    
errTrap:
    AssertNotSupportedException Err
End Sub

Public Sub GetAllInvariantCultures_WhenCalled_ContainsInvariantCulture()
    Dim Cultures() As CultureInfo
    Cultures = CultureInfo.GetCultures(CultureTypes.AllCultures)
    
    Dim Invariant As CultureInfo
    Set Invariant = CultureInfo.InvariantCulture
    
    Assert.That Cultures, Has.Some.EqualTo(Invariant).Using(New ObjectComparer)
End Sub

Public Sub OptionalCalendars_WhenCalled_DoesNotReturnNull()
    Dim Item As Variant
    For Each Item In CultureInfo.GetCultures(AllCultures)
        Dim Culture As CultureInfo
        Set Culture = Item
        
        Dim Actual() As Calendar
        Actual = Culture.OptionalCalendars
        Assert.That cArray.IsNull(Actual), Iz.False, Culture.LCID & " " & Culture.Name
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
    Assert.That c.DateTimeFormat.LongDatePattern, Iz.EqualTo("dddd, MMMM d, yyyy")
    Assert.That c.DateTimeFormat.LongTimePattern, Iz.EqualTo("h:mm:ss tt")
End Sub

Public Sub Calendary_WithenUsCulture_ReturnsGegorian()
    Dim c As CultureInfo
    Set c = Cor.NewCultureInfo("en-US")
    
    Dim Actual As Calendar
    Set Actual = c.Calendar
    
    Assert.That TypeOf Actual Is GregorianCalendar, Iz.True
End Sub


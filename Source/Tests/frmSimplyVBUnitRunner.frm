VERSION 5.00
Object = "{7983BD3B-752A-43EA-9BFF-444BBA1FC293}#4.0#0"; "SimplyVBUnit.Component.ocx"
Begin VB.Form frmSimplyVBUnitRunner 
   Caption         =   "Simply VB Unit"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10770
   Icon            =   "frmSimplyVBUnitRunner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin SimplyVBComp.UIRunner UIRunner1 
      Height          =   5175
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSimplyVBUnitRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Namespaces available:
'       Assert.*            ie. Assert.IsTrue Value

' Adding a testcase:
'   Use AddTest <object>

' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a public sub named Setup if you want Setup run before each test in the class.
'   4. (Optionally) Add a public sub named Teardown if you want Teardown run after each test in the class.
'   5. Add public subs of the tests you want run. No parameters.

Private Sub Form_Load()
    SkipUnsupportedTimeZone = False

    AddMicrosoftWin32
    AddSystem
    AddSystemCollections
    AddSystemSecurityCryptography
    AddSystemResources
    AddSystemIO
    AddSystemText
    AddSystemGlobalization
    AddSystemNumerics
    AddSystemSecurity
    AddSystemDiagnostics
    AddSystemThreading
End Sub

Private Sub AddSystemSecurity()
    With Sim.NewTestSuite("System.Security")
        .Add New TestSecurityElement
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemDiagnostics()
    With Sim.NewTestSuite("System.Diagnostics")
        .Add New StopWatchTests
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemThreading()
    With Sim.NewTestSuite("System.Threading")
        .Add New TestTicker
        
        AddTest .This
    End With
End Sub

Private Sub AddSystem()
    With Sim.NewTestSuite("System")
        .Add New ExceptionTests
        .Add New SystemExceptionTests
        .Add New ArgumentExceptionTests
        .Add New ArgumentNullExceptionTests
        .Add New ArgumentOutOfRangeTests
        .Add New ExceptionMethodsTests
        .Add New InvalidCastExceptionTests
        .Add New CorStringTests
        .Add New CharTests
        .Add New CharEnumeratorTests
        .Add New BufferTests
        .Add New VersionTests
        .Add New TestRandom
        .Add New MathExtTests
        .Add New GuidTests
        .Add New BitConverterTests
        .Add New CorDateTimeTests
        .Add New TestEnvironment
        .Add New TimeZoneTests
        .Add New TimeSpanTests
        .Add New ArrayConstructorTests
        
        .Add NewSuite("Object", _
            New ObjectBaseTests, _
            New ObjectTests, _
            New ObjectToStringWithDoublesTests, _
            New ObjectToStringWithLongsTests, _
            New ObjectToStringWithSinglesTests, _
            New ObjectToStringWithValuesTests)
        
        .Add NewSuite("Convert", _
            New ToBase64Tests, _
            New FromBase64Tests)
    
        .Add NewSuite("CorArray", _
            New CorArrayTests, _
            New CorArrayCopyTests, _
            New CorArraySortTests, _
            New CorArrayBinarySearchTests, _
            New CorArrayReverseTests, _
            New CorArrayIndexOfTests, _
            New CorArrayLastIndexOfTests, _
            New CorArrayCreateInstanceTests, _
            New CorArrayFindTests)
    
        AddTest .This
    End With
End Sub

Private Sub AddSystemCollections()
    With Sim.NewTestSuite("System.Collections")
        .Add New BitArrayTests
        .Add New BitArrayEnumeratorTests
        .Add New ComparerTests
        .Add New CaseInsensitiveComparerTests

        .Add NewSuite("SortedList", _
            New SortedListTests, _
            New SortedListEnumeratorTests, _
            New SortedKeyListTests, _
            New SortedValueListTests)

        .Add NewSuite("ArrayList", _
            New ArrayListTests, _
            New ArrayListAdapterTests, _
            New ArrayListRangedTests, _
            New ArrayListRepeatTests, _
            New ArrayListEnumeratorTests, _
            New ReadOnlyArrayListTests, _
            New FixedSizeArrayListTests)
        
        .Add NewSuite("Queue", _
            New QueueTests, _
            New QueueEnumeratorTests)
                
        .Add NewSuite("Stack", _
            New StackTests, _
            New StackEnumeratorTests)
        
        .Add NewSuite("Hashtable", _
            New HashtableTests, _
            New HashtableEnumeratorTests, _
            New HashtableKeyCollectionTests, _
            New HashtableValueCollectionTests, _
            New DictionaryEntryTests)
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemSecurityCryptography()
    With Sim.NewTestSuite("System.Security.Cryptography")
        .Add New TestRNGCryptoServiceProvider
        .Add New TestToBase64Transform
        .Add New TestFromBase64Transform
        .Add New TestSHA1CryptoServiceProvider
        .Add New TestSHA1Managed
        .Add New TestSHA256Managed
        .Add New TestSHA512Managed
        .Add New TestSHA384Managed
        .Add New TestMD5CryptoServiceProvider
        .Add New TestRIPEMD160Managed
        .Add New TestHMACSHA1
        .Add New TestHMACSHA1Managed
        .Add New TestHMACSHA256
        .Add New TestHMACSHA384
        .Add New TestHMACSHA512
        .Add New TestHMACMD5
        .Add New TestHMACRIPEMD160
        .Add New TestMACTripleDES
        .Add New TestRfc2898DeriveBytes
        .Add New CryptoConfigTests
        .Add New TestCspParameters
        .Add New TestCspKeyContainerInfo
        .Add New TestRSACryptoServiceProvider
        .Add New TestRSASignAndVerify
        .Add New TestDSACryptoServiceProvider
        
        .Add NewSuite("CryptoStream", _
            New TestCryptoStream, _
            New TestCryptoStreamReadBase64, _
            New TestCryptoStreamWriteBase64, _
            New TestCryptoStreamFullBase64, _
            New TestCryptoStreamMultiBlock)
            
        .Add NewSuite("DESCryptoServiceProvider", _
            New TestDESWeakKeys, _
            New TestDESCryptoServiceProvider, _
            New TestDESEncryption, _
            New TestDESPaddingModes, _
            New TestDESDecryption)
        
        .Add NewSuite("TripleDESCryptoServiceProvider", _
            New TestTripleDESWeakKeys, _
            New TestTripleDESCryptoServiceProvider, _
            New TestTripleDESEncryption, _
            New TestTripleDESDecryption)
        
        .Add NewSuite("SymmetricalAlgorithm", _
            New TestSymmetricAlgorithmBase, _
            New TestSymmetricAlgorithmBaseKey)
        
        .Add NewSuite("RC2CryptoServiceProvider", _
            New TestRC2Encryption, _
            New TestRC2Decryption)
        
        .Add NewSuite("RijndaelManaged", _
            New TestRijndaelManaged128, _
            New TestRijndaelEncryptionECB, _
            New TestRijndaelEncryptionCBC, _
            New TestRijndaelEncryptionCFB, _
            New TestRijndaelDecryptionECB, _
            New TestRijndaelDecryptionCBC, _
            New TestRijndaelDecryptionCFB)
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemResources()
    With Sim.NewTestSuite("System.Resources")
        .Add New TestResourceKey
        .Add New TestResourceWriter
        .Add New TestResourceManager
        .Add New TestResourceSet
        .Add New TestResourceReader
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemIO()
    With Sim.NewTestSuite("System.IO")
        .Add New BinaryReaderTests
        .Add New BinaryWriterTests
        .Add New FileInfoTests
        .Add New FileTests
        .Add New StreamReaderTests
        .Add New TestMappedFile
        .Add New FileNotFoundExceptionTests
        .Add New TestINIFile
        .Add New DriveInfoTests
        .Add New StringReaderTests
        .Add New StringWriterTests
        .Add New DirectoryTests
        .Add New DirectoryInfoTests
        .Add New MemoryStreamTests
        .Add New PathTests
        .Add New FileStreamTests
        .Add New StreamWriterTests
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemText()
    With Sim.NewTestSuite("System.Text")
        .Add New EncodingArgumentTests
        .Add New ASCIIEncodingTests
        .Add New TestUnicodeEncodingBig
        .Add New TestUnicodeEncoding
        .Add New TestDetermineEncoding
        .Add New TestEncoding437
        .Add New StringBuilderTests
        .Add New TestCustomFormatter
        
        .Add NewSuite("UTF7Encoding", _
            New TestUTF7GetChars, _
            New TestUTF7GetCharCount, _
            New TestUTF7GetBytes, _
            New TestUTF7GetByteCount)
        
        .Add NewSuite("UTF8Encoding", _
            New TestUTF8GetChars, _
            New TestUTF8GetCharCount, _
            New TestUTF8Encoding, _
            New TestUTF8GetByteCount)
        
        AddTest .This
    End With
End Sub

Private Sub AddMicrosoftWin32()
    With Sim.NewTestSuite("Microsoft.Win32")
        .Add New SafeHandleTests
        
        .Add NewSuite("RegistryKey", _
            New TestRegistryDeleteValue, _
            New TestRegistryKeySetGetValue, _
            New TestRegistryRootKeys, _
            New TestRegistryKey, _
            New TestRegistrySetValues)
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemGlobalization()
    With Sim.NewTestSuite("System.Globalization")
        .Add New TestThaiBuddhistCalendar
        .Add New TestTaiwanCalendar
        .Add New TestKoreanCalendar
        .Add New TestJapaneseCalendar
        .Add New TestHebrewCalendar
        .Add New TestJulianCalendar
        .Add New TestCodePageDecoder
        .Add New TestGregorianCalendar
        .Add New TestHijriCalendar
        .Add New CultureInfoTests
        .Add New DateTimeFormatInfoTests
        
        AddTest .This
    End With
End Sub

Private Sub AddSystemNumerics()
    With Sim.NewTestSuite("System.Numerics")
        .Add NewSuite("BigInteger", _
            New VBAdditionTests, _
            New VBBitTests, _
            New VBComparisonTests, _
            New VBCreateFromArraysTests, _
            New VBCreateFromNumbersTests, _
            New VBDivisionTests, _
            New VBFactorialTests, _
            New VBMultiplyTests, _
            New VBParseBinaryTests, _
            New VBParseDecimalTests, _
            New VBParseHexTests, _
            New VBPowTests, _
            New VBRightShiftTests, _
            New VBRndTests, _
            New VBShiftLeftTests, _
            New VBSquareRootTests, _
            New VBSubtractionTests, _
            New VBToBinaryStringTests, _
            New VBToStringDecimalTests, _
            New VBToStringHexTests, _
            New VBUnaryTests)
        
        AddTest .This
    End With
End Sub

Private Function NewSuite(ByVal Name As String, ParamArray Fixtures() As Variant) As TestSuite
    Dim Suite As TestSuite
    Set Suite = Sim.NewTestSuite(Name)
    
    Dim Fixture As Variant
    For Each Fixture In Fixtures
        Suite.Add Fixture
    Next
    
    Set NewSuite = Suite
End Function

Private Sub Form_Initialize()
    Me.UIRunner1.Init App
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF5
            UIRunner1.Run
    End Select
End Sub



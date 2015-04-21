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
    Dim Suite As TestSuite
    
    ' Add test cases here.
    '
    'AddTest New MyTestCase
    Const CATEGORY_CRYPTOGRAPHY As String = "Cryptography Tests"
    
    Dim System As TestSuite
    Set System = Sim.NewTestSuite("System")
    
    With Sim.NewTestSuite("Object")
        .Add New ObjectBaseTests
        .Add New ObjectTests
        .Add New ObjectToStringWithDoublesTests
        .Add New ObjectToStringWithLongsTests
        .Add New ObjectToStringWithSinglesTests
        .Add New ObjectToStringWithValuesTests
        
        System.Add .This
    End With
    
    System.Add New ExceptionTests
    System.Add New TestSystemException
    System.Add New TestDefaultSystemEx
    System.Add New ArgumentExceptionTests
    System.Add New ArgumentNullExceptionTests
    System.Add New ArgumentOutOfRangeTests
    System.Add New TestExceptionMethods
    System.Add New TestInvalidCastException
    System.Add New TestDefInvalidCast
    System.Add New CStringTests
    System.Add New CharTests
    System.Add New BufferTests
    System.Add New TestVersion
    System.Add New TestRandom
    System.Add New TestMathExt
    System.Add New GuidTests
    System.Add New BitConverterTests
    System.Add New CDateTimeTests
    System.Add New TestEnvironment
    System.Add New TimeZoneTests

    System.Add NewSuite("Convert", New ToBase64Tests, New FromBase64Tests)
    
    Dim cArrayTests As TestSuite
    Set cArrayTests = Sim.NewTestSuite("Array")
    cArrayTests.Add New cArrayTests
    cArrayTests.Add New cArrayCopyTests
    cArrayTests.Add New cArraySortTests
    cArrayTests.Add New cArrayBinarySearchTests
    cArrayTests.Add New cArrayReverseTests
    cArrayTests.Add New cArrayIndexOfTests
    cArrayTests.Add New cArrayLastIndexOfTests
    cArrayTests.Add New cArrayCreateInstanceTests
    cArrayTests.Add New cArrayFindTests
    System.Add cArrayTests
    System.Add New TimeSpanTests
    
    AddTest System
    
    AddTest CreateCollectionsTests
    
    Dim Cyrptography As TestSuite
    Set Cyrptography = Sim.NewTestSuite("System.Security.Cryptography")
    Set Cyrptography.Categories = Sim.NewCategorization(CATEGORY_CRYPTOGRAPHY, True)
    Cyrptography.Categories.Add CATEGORY_CRYPTOGRAPHY
    Cyrptography.Add New TestRNGCryptoServiceProvider
    Cyrptography.Add New TestToBase64Transform
    Cyrptography.Add New TestFromBase64Transform
    Cyrptography.Add NewSuite("CryptoStream", New TestCryptoStream, New TestCryptoStreamReadBase64, New TestCryptoStreamWriteBase64, New TestCryptoStreamFullBase64, New TestCryptoStreamMultiBlock)
    Cyrptography.Add NewSuite("DESCryptoServiceProvider", New TestDESWeakKeys, New TestDESCryptoServiceProvider, New TestDESEncryption, New TestDESPaddingModes, New TestDESDecryption)
    Cyrptography.Add NewSuite("TripleDESCryptoServiceProvider", New TestTripleDESWeakKeys, New TestTripleDESCryptoServiceProvider, New TestTripleDESEncryption, New TestTripleDESDecryption)
    Cyrptography.Add NewSuite("SymmetricalAlgorithm", New TestSymmetricAlgorithmBase, New TestSymmetricAlgorithmBaseKey)
    Cyrptography.Add NewSuite("RC2CryptoServiceProvider", New TestRC2Encryption, New TestRC2Decryption)
    Cyrptography.Add NewSuite("RijndaelManaged", _
                              New TestRijndaelManaged128, _
                              New TestRijndaelEncryptionECB, _
                              New TestRijndaelEncryptionCBC, _
                              New TestRijndaelEncryptionCFB, _
                              New TestRijndaelDecryptionECB, _
                              New TestRijndaelDecryptionCBC, _
                              New TestRijndaelDecryptionCFB)
    Dim HashTests As TestSuite
    Set HashTests = Sim.NewTestSuite("Hash Tests")
    HashTests.Add New TestSHA1CryptoServiceProvider
    HashTests.Add New TestSHA1Managed
    HashTests.Add New TestSHA256Managed
    HashTests.Add New TestSHA512Managed
    HashTests.Add New TestSHA384Managed
    HashTests.Add New TestMD5CryptoServiceProvider
    HashTests.Add New TestRIPEMD160Managed
    Cyrptography.Add HashTests
    Dim HMACTests As TestSuite
    Set HMACTests = Sim.NewTestSuite("HMAC Tests")
    HMACTests.Add New TestHMACSHA1
    HMACTests.Add New TestHMACSHA1Managed
    HMACTests.Add New TestHMACSHA256
    HMACTests.Add New TestHMACSHA384
    HMACTests.Add New TestHMACSHA512
    HMACTests.Add New TestHMACMD5
    HMACTests.Add New TestHMACRIPEMD160
    Cyrptography.Add HMACTests
    Cyrptography.Add New TestMACTripleDES
    Cyrptography.Add New TestRfc2898DeriveBytes
    Cyrptography.Add New CryptoConfigTests
    Dim RSATests As TestSuite
    Set RSATests = Sim.NewTestSuite("RSACryptoServiceProvider")
    RSATests.Add New TestCspParameters
    RSATests.Add New TestCspKeyContainerInfo
    RSATests.Add New TestRSACryptoServiceProvider
    RSATests.Add New TestRSASignAndVerify
    Cyrptography.Add RSATests
    Cyrptography.Add New TestDSACryptoServiceProvider
    AddTest Cyrptography
    
    AddTest NewSuite("System.Security", New TestSecurityElement)
    AddTest NewSuite("System.Diagnostics", New StopWatchTests)
    
    Dim Resources As TestSuite
    Set Resources = Sim.NewTestSuite("System.Resources")
    Resources.Add New TestResourceKey
    Resources.Add New TestResourceWriter
    Resources.Add New TestResourceManager
    Resources.Add New TestResourceSet
    Resources.Add New TestResourceReader
    'AddTest New TestWinResourceReader
    
    AddTest Resources
    
    AddTest NewSuite("System.Threading", New TestTicker)
    
    Dim IO As TestSuite
    Set IO = Sim.NewTestSuite("System.IO")
    IO.Add New BinaryReaderTests
    IO.Add New BinaryWriterTests
    IO.Add New TestFileInfo
    IO.Add New FileTests
    IO.Add New TestStreamReader
    IO.Add New TestMappedFile
    IO.Add New TestFileNotFoundException
    IO.Add New TestINIFile
    IO.Add New DriveInfoTests
    IO.Add New TestStringReader
    IO.Add New TestStringWriter
    IO.Add New DirectoryTests
    IO.Add New DirectoryInfoTests
    IO.Add New MemoryStreamTests
    IO.Add New PathTests

    Dim StreamWriterTests As TestSuite
    Set StreamWriterTests = Sim.NewTestSuite("StreamWriter")
    StreamWriterTests.Add New TestStreamWriter
    StreamWriterTests.Add New TestStreamWriterWithMem
    StreamWriterTests.Add New TestSWWithMemAutoFlush
    IO.Add StreamWriterTests

    Dim FileStreamTests As TestSuite
    Set FileStreamTests = Sim.NewTestSuite("FileStream")
'    FileStreamTests.Add New TestFileStreamWrite
'    FileStreamTests.Add New TestFileStreamSmallBuffer
    FileStreamTests.Add New FileStreamTests
    IO.Add FileStreamTests

    AddTest IO
    
    Dim Text As TestSuite
    Set Text = Sim.NewTestSuite("System.Text")
    Text.Add New EncodingArgumentTests
    Text.Add New ASCIIEncodingTests
    Text.Add New TestUnicodeEncodingBig
    Text.Add New TestUnicodeEncoding
    Text.Add New TestDetermineEncoding
    Text.Add New TestEncoding437
    Text.Add New StringBuilderTests
    Text.Add New TestCustomFormatter

    Dim UTF7EncodingTests As TestSuite
    Set UTF7EncodingTests = Sim.NewTestSuite("UTF7Encoding")
    UTF7EncodingTests.Add New TestUTF7GetChars
    UTF7EncodingTests.Add New TestUTF7GetCharCount
    UTF7EncodingTests.Add New TestUTF7GetBytes
    UTF7EncodingTests.Add New TestUTF7GetByteCount
    Text.Add UTF7EncodingTests
    
    Dim UTF8EncodingTests As TestSuite
    Set UTF8EncodingTests = Sim.NewTestSuite("UTF8Encoding")
    UTF8EncodingTests.Add New TestUTF8GetChars
    UTF8EncodingTests.Add New TestUTF8GetCharCount
    UTF8EncodingTests.Add New TestUTF8Encoding
    UTF8EncodingTests.Add New TestUTF8GetByteCount
    Text.Add UTF8EncodingTests
    
    AddTest Text
    
    Dim Win32 As TestSuite
    Set Win32 = Sim.NewTestSuite("Microsoft.Win32")
    Dim RegistryKeyTests As TestSuite
    Set RegistryKeyTests = Sim.NewTestSuite("RegistryKey")
    RegistryKeyTests.Add New TestRegistryDeleteValue
    RegistryKeyTests.Add New TestRegistryKeySetGetValue
    RegistryKeyTests.Add New TestRegistryRootKeys
    RegistryKeyTests.Add New TestRegistryKey
    RegistryKeyTests.Add New TestRegistrySetValues
    Win32.Add RegistryKeyTests
    Win32.Add New SafeHandleTests
    AddTest Win32
    
    Dim Globalization As TestSuite
    Set Globalization = Sim.NewTestSuite("System.Globalization")
    Globalization.Add New TestThaiBuddhistCalendar
    Globalization.Add New TestTaiwanCalendar
    Globalization.Add New TestKoreanCalendar
    Globalization.Add New TestJapaneseCalendar
    Globalization.Add New TestHebrewCalendar
    Globalization.Add New TestJulianCalendar
    Globalization.Add New TestCodePageDecoder
    Globalization.Add New CharEnumeratorTests
    Globalization.Add New TestGregorianCalendar
    Globalization.Add New TestHijriCalendar
    Globalization.Add New CultureInfoTests
    Globalization.Add New TestDateTimeFormatInfoInv

    AddTest Globalization
        
    'AddTest New TestWeakReference
    
    Dim Numerics As TestSuite
    Set Numerics = Sim.NewTestSuite("System.Numerics")
    Dim BigIntegerTests As TestSuite
    Set BigIntegerTests = Sim.NewTestSuite("BigInteger")
    BigIntegerTests.Add New VBAdditionTests
    BigIntegerTests.Add New VBBitTests
    BigIntegerTests.Add New VBComparisonTests
    BigIntegerTests.Add New VBCreateFromArraysTests
    BigIntegerTests.Add New VBCreateFromNumbersTests
    BigIntegerTests.Add New VBDivisionTests
    BigIntegerTests.Add New VBFactorialTests
    BigIntegerTests.Add New VBMultiplyTests
    BigIntegerTests.Add New VBParseBinaryTests
    BigIntegerTests.Add New VBParseDecimalTests
    BigIntegerTests.Add New VBParseHexTests
    BigIntegerTests.Add New VBPowTests
    BigIntegerTests.Add New VBRightShiftTests
    BigIntegerTests.Add New VBRndTests
    BigIntegerTests.Add New VBShiftLeftTests
    BigIntegerTests.Add New VBSquareRootTests
    BigIntegerTests.Add New VBSubtractionTests
    BigIntegerTests.Add New VBToBinaryStringTests
    BigIntegerTests.Add New VBToStringDecimalTests
    BigIntegerTests.Add New VBToStringHexTests
    BigIntegerTests.Add New VBUnaryTests
    Numerics.Add BigIntegerTests
    AddTest Numerics
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



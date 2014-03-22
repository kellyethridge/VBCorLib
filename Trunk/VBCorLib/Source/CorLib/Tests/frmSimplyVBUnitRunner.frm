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
    
    ' Add test cases here.
    '
    'AddTest New MyTestCase
    Const CATEGORY_CRYPTOGRAPHY As String = "Cryptography Tests"
        
    Dim CryptoTests As TestSuite
    Set CryptoTests = Sim.NewTestSuite("Cryptography Tests")
    Set CryptoTests.Categories = Sim.NewCategorization("Cryptography Tests", True)
    CryptoTests.Categories.Add CATEGORY_CRYPTOGRAPHY
    CryptoTests.Add New TestRNGCryptoServiceProvider
    CryptoTests.Add New TestToBase64Transform
    CryptoTests.Add New TestFromBase64Transform
    CryptoTests.Add New TestCryptoStream
    CryptoTests.Add New TestCryptoStreamReadBase64
    CryptoTests.Add New TestCryptoStreamWriteBase64
    CryptoTests.Add New TestCryptoStreamFullBase64
    CryptoTests.Add New TestDESWeakKeys
    CryptoTests.Add New TestDESCryptoServiceProvider
    CryptoTests.Add New TestDESEncryption
    CryptoTests.Add New TestDESPaddingModes
    CryptoTests.Add New TestDESDecryption
    CryptoTests.Add New TestCryptoStreamMultiBlock
    CryptoTests.Add New TestTripleDESWeakKeys
    CryptoTests.Add New TestTripleDESCryptoServiceProvider
    CryptoTests.Add New TestSymmetricAlgorithmBase
    CryptoTests.Add New TestSymmetricAlgorithmBaseKey
    CryptoTests.Add New TestRC2Encryption
    CryptoTests.Add New TestRC2Decryption
    CryptoTests.Add New TestTripleDESDecryption
    CryptoTests.Add New TestTripleDESEncryption
    
    Dim RijndaelTests As TestSuite
    Set RijndaelTests = Sim.NewTestSuite("Rijndael Tests")
    RijndaelTests.Add New TestRijndaelManaged128
    RijndaelTests.Add New TestRijndaelEncryptionECB
    RijndaelTests.Add New TestRijndaelEncryptionCBC
    RijndaelTests.Add New TestRijndaelEncryptionCFB
    RijndaelTests.Add New TestRijndaelDecryptionECB
    RijndaelTests.Add New TestRijndaelDecryptionCBC
    RijndaelTests.Add New TestRijndaelDecryptionCFB
    CryptoTests.Add RijndaelTests
    
    Dim HashTests As TestSuite
    Set HashTests = Sim.NewTestSuite("Hash Tests")
    HashTests.Add New TestSHA1CryptoServiceProvider
    HashTests.Add New TestSHA1Managed
    HashTests.Add New TestSHA256Managed
    HashTests.Add New TestSHA512Managed
    HashTests.Add New TestSHA384Managed
    HashTests.Add New TestMD5CryptoServiceProvider
    HashTests.Add New TestRIPEMD160Managed
    CryptoTests.Add HashTests
    
    Dim HMACTests As TestSuite
    Set HMACTests = Sim.NewTestSuite("HMAC Tests")
    HMACTests.Add New TestHMACSHA1
    HMACTests.Add New TestHMACSHA1Managed
    HMACTests.Add New TestHMACSHA256
    HMACTests.Add New TestHMACSHA384
    HMACTests.Add New TestHMACSHA512
    HMACTests.Add New TestHMACMD5
    HMACTests.Add New TestHMACRIPEMD160
    
    CryptoTests.Add HMACTests
    CryptoTests.Add New TestMACTripleDES
    CryptoTests.Add New TestRfc2898DeriveBytes
    CryptoTests.Add New TestCryptoConfig
    
    Dim RSATests As TestSuite
    Set RSATests = Sim.NewTestSuite("RSA Tests")
    RSATests.Add New TestCspParameters
    RSATests.Add New TestCspKeyContainerInfo
    RSATests.Add New TestRSACryptoServiceProvider
    RSATests.Add New TestRSASignAndVerify
    CryptoTests.Add RSATests
    
    CryptoTests.Add New TestDSACryptoServiceProvider
    
    AddTest CryptoTests
    
    AddTest New TestSecurityElement

    AddTest New TestStopWatch
    AddTest New TestResourceKey
    'AddTest New TestWinResourceReader
    AddTest New TestResourceWriter
    AddTest New TestTicker
    AddTest New TestINIFile
    AddTest New TestDriveInfo
    AddTest New TestCustomFormatter
    AddTest New TestResourceManager
    AddTest New TestHashTableHCP
    AddTest New TestResourceSet
    AddTest New TestCaseInsensitiveHCP
    AddTest New TestResourceReader
    
'    AddTest New ConvertTests
    Dim ConvertTests As TestSuite
    Set ConvertTests = Sim.NewTestSuite("Convert Tests")
    ConvertTests.Add New ToBase64Tests
    ConvertTests.Add New FromBase64Tests
    ConvertTests.Add New ToStringWithLongsTests
    ConvertTests.Add New ToStringWithDoublesTests
    ConvertTests.Add New ToStringWithSinglesTests

    ConvertTests.Add New ConvertTests
    AddTest ConvertTests
    
    AddTest New TestMathExt
    AddTest New TestGuid
    AddTest New TestASCIIEncoding
    AddTest New TestHijriCalendar
    
    Dim RegistryKeyTests As TestSuite
    Set RegistryKeyTests = Sim.NewTestSuite("RegistryKey Tests")
    RegistryKeyTests.Add New TestRegistryDeleteValue
    RegistryKeyTests.Add New TestRegistryKeySetGetValue
    RegistryKeyTests.Add New TestRegistryRootKeys
    RegistryKeyTests.Add New TestRegistryKey
    RegistryKeyTests.Add New TestRegistrySetValues
    AddTest RegistryKeyTests
    
    
    AddTest New TestThaiBuddhistCalendar
    AddTest New TestTaiwanCalendar
    AddTest New TestKoreanCalendar
    AddTest New TestJapaneseCalendar
    AddTest New TestHebrewCalendar
    AddTest New TestJulianCalendar
    AddTest New TestCodePageDecoder
    AddTest New TestEncoding437
    AddTest New TestCharEnumerator
    AddTest New TestGregorianCalendar
    AddTest New TestDetermineEncoding
    AddTest New TestBinaryReader
    AddTest New TestBinaryWriter
    AddTest New TestFileInfo
    AddTest New TestFile
    AddTest New TestStreamReader
    
    Dim StreamWriterTests As TestSuite
    Set StreamWriterTests = Sim.NewTestSuite("StreamWriter Tests")
    StreamWriterTests.Add New TestStreamWriter
    StreamWriterTests.Add New TestStreamWriterWithMem
    StreamWriterTests.Add New TestSWWithMemAutoFlush
    AddTest StreamWriterTests
        
    AddTest New TestDirectory
    AddTest New TestDirectoryInfo
    AddTest New TestStringReader
    AddTest New TestStringWriter
    AddTest New TestUnicodeEncodingBig
    AddTest New TestUnicodeEncoding
    
    Dim FileStreamTests As TestSuite
    Set FileStreamTests = Sim.NewTestSuite("FileStream Tests")
    FileStreamTests.Add New TestFileStreamWrite
    FileStreamTests.Add New TestFileStreamSmallBuffer
    FileStreamTests.Add New TestFileStream
    AddTest FileStreamTests
    
    Dim MemoryStreamTests As TestSuite
    Set MemoryStreamTests = Sim.NewTestSuite("MemoryStream Tests")
    MemoryStreamTests.Add New TestUserMemoryStream
    MemoryStreamTests.Add New TestMemoryStream
    AddTest MemoryStreamTests
    
    Dim UTF7EncodingTests As TestSuite
    Set UTF7EncodingTests = Sim.NewTestSuite("UTF7Encoding Tests")
    UTF7EncodingTests.Add New TestUTF7GetChars
    UTF7EncodingTests.Add New TestUTF7GetCharCount
    UTF7EncodingTests.Add New TestUTF7GetBytes
    UTF7EncodingTests.Add New TestUTF7GetByteCount
    AddTest UTF7EncodingTests
    
    Dim UTF8EncodingTests As TestSuite
    Set UTF8EncodingTests = Sim.NewTestSuite("UTF8Encoding Tests")
    UTF8EncodingTests.Add New TestUTF8GetChars
    UTF8EncodingTests.Add New TestUTF8GetCharCount
    UTF8EncodingTests.Add New TestUTF8Encoding
    UTF8EncodingTests.Add New TestUTF8GetByteCount
    AddTest UTF8EncodingTests
    
    AddTest New TestPath
    AddTest New TestEnvironment
    AddTest New TestTimeZone
    AddTest New TestDateTimeFormatInfoInv
    AddTest New TestCultureInfo
    AddTest New TestMappedFile
    AddTest New TestFileNotFoundException
    AddTest New TestcDateTime
    
    Dim TimeSpanTests As TestSuite
    Set TimeSpanTests = Sim.NewTestSuite("TimeSpan Tests")
    TimeSpanTests.Add New TestTimeSpan
    TimeSpanTests.Add New TestTimeSpan994394150ms
    TimeSpanTests.Add New TestTimeSpanCreation
    AddTest TimeSpanTests
    
    AddTest New TestVersion
    AddTest New TestRandom
    AddTest New TestBitConverter
    
    'AddTest New TestWeakReference
    AddTest New TestHashTable
    AddTest New TestBuffer
    AddTest New TestBitArray
    AddTest New TestSortedList
    AddTest New TestDictionaryEntry
    AddTest New TestQueue
    AddTest New TestStack
    
    Dim ArrayListTests As TestSuite
    Set ArrayListTests = Sim.NewTestSuite("ArrayList Tests")
    ArrayListTests.Add New TestArrayListExceptions
    ArrayListTests.Add New TestArrayListRange
    ArrayListTests.Add New TestArrayList10Items
    ArrayListTests.Add New TestArrayList
    ArrayListTests.Add New TestAdapterArrayList
    ArrayListTests.Add New TestAdapterArrayList10Items
    ArrayListTests.Add New TestAdapterBinarySearch
    AddTest ArrayListTests
    
    AddTest New TestcString
    AddTest New StringBuilderTests
    AddTest New TestDefaultComparer

    Dim ExceptionTests As TestSuite
    Set ExceptionTests = Sim.NewTestSuite("Exception Tests")
    ExceptionTests.Add New ExceptionTests
    ExceptionTests.Add New TestDefaultException
    ExceptionTests.Add New TestSystemException
    ExceptionTests.Add New TestDefaultSystemEx
    ExceptionTests.Add New TestArgumentException
    ExceptionTests.Add New TestDefaultArgumentEx
    ExceptionTests.Add New TestDefaultArgumentNull
    ExceptionTests.Add New TestArgumentNullException
    ExceptionTests.Add New TestArgumentOutOfRange
    ExceptionTests.Add New TestDefArgumentOutOfRange
    ExceptionTests.Add New TestExceptionMethods
    ExceptionTests.Add New TestInvalidCastException
    ExceptionTests.Add New TestDefInvalidCast
    AddTest ExceptionTests
    

    Dim cArrayTests As TestSuite
    Set cArrayTests = Sim.NewTestSuite("cArray Tests")
    cArrayTests.Add New cArrayTests
    cArrayTests.Add New cArrayCopyTests
    cArrayTests.Add New cArraySortTests
    cArrayTests.Add New cArrayBinarySearchTests
    cArrayTests.Add New cArrayReverseTests
    cArrayTests.Add New cArrayIndexOfTests
    cArrayTests.Add New cArrayLastIndexOfTests
    cArrayTests.Add New cArrayCreateInstanceTests
    cArrayTests.Add New cArrayFindTests
    AddTest cArrayTests
    
    
    Dim BigIntegerTests As TestSuite
    Set BigIntegerTests = Sim.NewTestSuite("BigInteger Tests")
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
    AddTest BigIntegerTests
    
    
End Sub



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



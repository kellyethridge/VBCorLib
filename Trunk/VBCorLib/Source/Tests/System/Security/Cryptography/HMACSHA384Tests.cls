VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACSHA384Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACSHA384Tests
'
' Test vector source: https://tools.ietf.org/html/rfc4231
'
Option Explicit
Implements ICategorizable
Implements ITestCaseSource

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Security.Cryptography"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ComputeHash_WithTextValues_ReturnsExpected"
            Test.Use "0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b", "Hi There", "afd03944d84895626b0825f4ab46907f15f9dadbe4101ec682aa034c7cebc59cfaea9ea9076ede7f4af152e8b2fa9cb6"
            Test.Use TextToHex("Jefe"), "what do ya want for nothing?", "af45d2e376484031617f78d2b58a6b1b9c7ef464f5a01b47e42ec3736322445e8e2240ca5e69e2c78b3239ecfab21649"
            Test.Use HexString(&HAA, 131), "Test Using Larger Than Block-Size Key - Hash Key First", "4ece084485813e9088d2c63a041bc5b44f9ef1012a2b588f3cd11f05033ac4c60c2ef6ab4030fe8296248df163f44952"
            Test.Use HexString(&HAA, 131), "This is a test using a larger than block-size key and a larger than block-size data. The key needs to be hashed before being used by the HMAC algorithm.", "6617178e941f020d351e2f254e8fd32c602420feb0b8fb9adccebb82461e99c5a678cc31e799176d3860e6110c46523e"
            
        Case "ComputeHash_WithHexValues_ReturnsExpected"
            Test.Use "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", HexString(&HDD, 50), "88062608d3e6ad8a0aa2ace014c8a86f0aa635d947ac9febe83ef4e55966144b2a5ab39dc13814b94e3ab6e101a34f27"
            Test.Use "0102030405060708090a0b0c0d0e0f10111213141516171819", HexString(&HCD, 50), "3e8a69b7783c25851933ab6290af6ca77a9981480850009cc5577c6e1f573b4e6801dd23c4a7d679ccf8a386c674cffb"
                        
    End Select
End Sub

Public Sub ComputeHash_WithTextValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim csp As HMAC
    Set csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Dim i As Long
    For i = LBound(Actual) To UBound(Actual)
        Debug.Print CorString.Format("{0:x2}", Actual(i));
    Next
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithHexValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim csp As HMAC
    Set csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(HexToBytes(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithSameValueOnMultipleCalls_ReturnsSameResult()
    Dim csp As HMAC
    Set csp = CreateHMAC
    Dim Source() As Byte
    Source = Encoding.UTF8.GetBytes("Hi there")
    
    Dim Hash1() As Byte
    Dim Hash2() As Byte
    Hash1 = csp.ComputeHash(Source)
    Hash2 = csp.ComputeHash(Source)
    
    Assert.That Hash1, Iz.EqualTo(Hash2)
End Sub

Public Sub Key_WithDefaultConstructor_ReturnsRandom64ByteArray()
    Dim csp As HMAC
    Set csp = CreateHMAC
    
    Dim Actual() As Byte
    Actual = csp.Key
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That CorArray.Length(Actual), Iz.EqualTo(64)
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub Constructor_WithNullKey_ThrowsArgumentNullException()
    Dim NullKey() As Byte
    
    On Error Resume Next
    Cor.NewHMACSHA384 NullKey
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub TransformBlock_WhenCalled_HashesValue()
    Dim csp As HMAC
    Dim Block1() As Byte
    Dim Block2() As Byte
    Set csp = CreateHMAC("0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b")
    Block1 = Encoding.UTF8.GetBytes("Hi ")
    Block2 = Encoding.UTF8.GetBytes("There")
    
    Dim Actual() As Byte
    csp.TransformBlock Block1, 0, 3, Block1, 0
    csp.TransformFinalBlock Block2, 0, 5
    Actual = csp.Hash
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("afd03944d84895626b0825f4ab46907f15f9dadbe4101ec682aa034c7cebc59cfaea9ea9076ede7f4af152e8b2fa9cb6"))
End Sub

Public Sub TransformFinalBlock_WhenCalledMultipleTimes_ReturnsSameResult()
    Dim csp As HMAC
    Set csp = CreateHMAC
    Dim Source() As Byte
    Source = Encoding.UTF8.GetBytes("Hi there")
    
    Dim Hash1() As Byte
    Dim Hash2() As Byte
    csp.TransformFinalBlock Source, 0, 8
    Hash1 = csp.Hash
    csp.TransformFinalBlock Source, 0, 8
    Hash2 = csp.Hash
    
    Assert.That Hash1, Iz.EqualTo(Hash2)
End Sub

Private Function CreateHMAC(Optional ByRef Key As String) As HMAC
    If Len(Key) > 0 Then
        Set CreateHMAC = Cor.NewHMACSHA384(HexToBytes(Key))
    Else
        Set CreateHMAC = New HMACSHA384
    End If
End Function



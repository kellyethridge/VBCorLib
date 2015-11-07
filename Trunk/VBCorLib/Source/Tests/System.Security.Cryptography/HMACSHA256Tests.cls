VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACSHA256Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACSHA256Tests
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
            Test.Use "0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b", "Hi There", "b0344c61d8db38535ca8afceaf0bf12b881dc200c9833da726e9376c2e32cff7"
            Test.Use TextToHex("Jefe"), "what do ya want for nothing?", "5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843"
            Test.Use HexString(&HAA, 131), "Test Using Larger Than Block-Size Key - Hash Key First", "60e431591ee0b67f0d8a26aacbf5b77f8e0bc6213728c5140546040f0ee37f54"
            Test.Use HexString(&HAA, 131), "This is a test using a larger than block-size key and a larger than block-size data. The key needs to be hashed before being used by the HMAC algorithm.", "9b09ffa71b942fcb27635fbcd5b0e944bfdc63644f0713938a7f51535c3a35e2"
            
        Case "ComputeHash_WithHexValues_ReturnsExpected"
            Test.Use "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", HexString(&HDD, 50), "773ea91e36800e46854db8ebd09181a72959098b3ef8c122d9635514ced565fe"
            Test.Use "0102030405060708090a0b0c0d0e0f10111213141516171819", HexString(&HCD, 50), "82558a389a443c0ea4cc819899f2083a85f0faa3e578f8077a2e3ff46729665b"
                        
    End Select
End Sub

Public Sub ComputeHash_WithTextValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim csp As HMAC
    Set csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
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
    Cor.NewHMACSHA256 NullKey
    
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
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("b0344c61d8db38535ca8afceaf0bf12b881dc200c9833da726e9376c2e32cff7"))
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
        Set CreateHMAC = Cor.NewHMACSHA256(HexToBytes(Key))
    Else
        Set CreateHMAC = New HMACSHA256
    End If
End Function


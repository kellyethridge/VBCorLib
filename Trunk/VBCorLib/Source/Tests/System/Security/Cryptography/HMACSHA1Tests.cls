VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACSHA1Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACSHA1Tests
'
' Test vector source: https://tools.ietf.org/html/rfc2202
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
            Test.Use "0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b", "Hi There", "b617318655057264e28bc0b6fb378c8ef146be00"
            Test.Use TextToHex("Jefe"), "what do ya want for nothing?", "effcdf6ae5eb2fa2d27416d5f184df9c259a7c79"
            Test.Use "0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c", "Test With Truncation", "4c1a03424b55e07fe7f27be1d58bb9324a9a5a04"
            Test.Use HexString(&HAA, 80), "Test Using Larger Than Block-Size Key - Hash Key First", "aa4ae5e15272d00e95705637ce8a3b55ed402112"
            Test.Use HexString(&HAA, 80), "Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data", "e8e99d0f45237d786d6bbaa7965c7808bbff1a91"
            
        Case "ComputeHash_WithHexValues_ReturnsExpected"
            Test.Use "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", HexString(&HDD, 50), "125d7342b9ac11cd91a39af48aa17b4f63f175d3"
            Test.Use "0102030405060708090a0b0c0d0e0f10111213141516171819", HexString(&HCD, 50), "4c9007f4026250c6bc8414f9bf50c86c2d7235da"
                        
    End Select
End Sub

Public Sub ComputeHash_WithTextValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim Csp As HMAC
    Set Csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = Csp.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithHexValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim Csp As HMAC
    Set Csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = Csp.ComputeHash(HexToBytes(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithSameValueOnMultipleCalls_ReturnsSameResult()
    Dim Csp As HMAC
    Set Csp = CreateHMAC
    Dim Source() As Byte
    Source = Encoding.UTF8.GetBytes("Hi there")
    
    Dim Hash1() As Byte
    Dim Hash2() As Byte
    Hash1 = Csp.ComputeHash(Source)
    Hash2 = Csp.ComputeHash(Source)
    
    Assert.That Hash1, Iz.EqualTo(Hash2)
End Sub

Public Sub Key_WithDefaultConstructor_ReturnsRandom64ByteArray()
    Dim Csp As HMAC
    Set Csp = CreateHMAC
    
    Dim Actual() As Byte
    Actual = Csp.Key
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That CorArray.Length(Actual), Iz.EqualTo(64)
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub Constructor_WithNullKey_ThrowsArgumentNullException()
    Dim NullKey() As Byte
    
    On Error Resume Next
    Cor.NewHMACSHA1 NullKey
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub TransformBlock_WhenCalled_HashesValue()
    Dim Csp As HMAC
    Dim Block1() As Byte
    Dim Block2() As Byte
    Set Csp = CreateHMAC("0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b")
    Block1 = Encoding.UTF8.GetBytes("Hi ")
    Block2 = Encoding.UTF8.GetBytes("There")
    
    Dim Actual() As Byte
    Csp.TransformBlock Block1, 0, 3, Block1, 0
    Csp.TransformFinalBlock Block2, 0, 5
    Actual = Csp.Hash
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("b617318655057264e28bc0b6fb378c8ef146be00"))
End Sub

Public Sub TransformFinalBlock_WhenCalledMultipleTimes_ReturnsSameResult()
    Dim Csp As HMAC
    Set Csp = CreateHMAC
    Dim Source() As Byte
    Source = Encoding.UTF8.GetBytes("Hi there")
    
    Dim Hash1() As Byte
    Dim Hash2() As Byte
    Csp.TransformFinalBlock Source, 0, 8
    Hash1 = Csp.Hash
    Csp.TransformFinalBlock Source, 0, 8
    Hash2 = Csp.Hash
    
    Assert.That Hash1, Iz.EqualTo(Hash2)
End Sub

Private Function CreateHMAC(Optional ByRef Key As String) As HMAC
    If Len(Key) > 0 Then
        Set CreateHMAC = Cor.NewHMACSHA1(HexToBytes(Key))
    Else
        Set CreateHMAC = New HMACSHA1
    End If
End Function

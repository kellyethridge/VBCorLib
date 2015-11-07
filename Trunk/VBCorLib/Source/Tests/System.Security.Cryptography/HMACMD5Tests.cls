VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACMD5Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACMD5Tests
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
            Test.Use "0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b", "Hi There", "9294727a3638bb1c13f48ef8158bfc9d"
            Test.Use TextToHex("Jefe"), "what do ya want for nothing?", "750c783e6ab0b503eaa86e310a5db738"
            Test.Use "0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c", "Test With Truncation", "56461ef2342edc00f9bab995690efd4c"
            Test.Use HexString(&HAA, 80), "Test Using Larger Than Block-Size Key - Hash Key First", "6b1ab7fe4bd7bf8f0b62e6ce61b9d0cd"
            Test.Use HexString(&HAA, 80), "Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data", "6f630fad67cda0ee1fb1f562db3aa53e"
            
        Case "ComputeHash_WithHexValues_ReturnsExpected"
            Test.Use "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", HexString(&HDD, 50), "56be34521d144c88dbb8c733f0e8b3f6"
            Test.Use "0102030405060708090a0b0c0d0e0f10111213141516171819", HexString(&HCD, 50), "697eaf0aca3a3aea3a75164746ffaa79"
                        
    End Select
End Sub

Public Sub ComputeHash_WithTextValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim Csp As HMACMD5
    Set Csp = Cor.NewHMACMD5(HexToBytes(Key))
    
    Dim Actual() As Byte
    Actual = Csp.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithHexValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim Csp As HMACMD5
    Set Csp = Cor.NewHMACMD5(HexToBytes(Key))
    
    Dim Actual() As Byte
    Actual = Csp.ComputeHash(HexToBytes(Message))
    
    Assert.That Actual, Iz.EqualTo(HexToBytes(Expected))
End Sub

Public Sub ComputeHash_WithSameValueOnMultipleCalls_ReturnsSameResult()
    Dim Csp As New HMACMD5
    Dim Source() As Byte
    Source = Encoding.UTF8.GetBytes("Hi there")
    
    Dim Hash1() As Byte
    Dim Hash2() As Byte
    Hash1 = Csp.ComputeHash(Source)
    Hash2 = Csp.ComputeHash(Source)
    
    Assert.That Hash1, Iz.EqualTo(Hash2)
End Sub

Public Sub Key_WithDefaultConstructor_ReturnsRandom64ByteArray()
    Dim Csp As New HMACMD5
    
    Dim Actual() As Byte
    Actual = Csp.Key
    
    Assert.That CorArray.IsNull(Actual), Iz.False
    Assert.That CorArray.Length(Actual), Iz.EqualTo(64)
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub Constructor_WithNullKey_ThrowsArgumentNullException()
    Dim NullKey() As Byte
    
    On Error Resume Next
    Cor.NewHMACMD5 NullKey
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub TransformBlock_WhenCalled_HashesValue()
    Dim Csp As HMACMD5
    Dim Block1() As Byte
    Dim Block2() As Byte
    Set Csp = Cor.NewHMACMD5(HexToBytes("0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b"))
    Block1 = Encoding.UTF8.GetBytes("Hi ")
    Block2 = Encoding.UTF8.GetBytes("There")
    
    Dim Actual() As Byte
    Csp.TransformBlock Block1, 0, 3, Block1, 0
    Csp.TransformFinalBlock Block2, 0, 5
    Actual = Csp.Hash
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("9294727a3638bb1c13f48ef8158bfc9d"))
End Sub

Public Sub TransformFinalBlock_WhenCalledMultipleTimes_ReturnsSameResult()
    Dim Csp As New HMACMD5
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



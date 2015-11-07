VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACSHA512Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACSHA512Tests
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
            Test.Use "0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b", "Hi There", "87aa7cdea5ef619d4ff0b4241a1d6cb02379f4e2ce4ec2787ad0b30545e17cdedaa833b7d6b8a702038b274eaea3f4e4be9d914eeb61f1702e696c203a126854"
            Test.Use TextToHex("Jefe"), "what do ya want for nothing?", "164b7a7bfcf819e2e395fbe73b56e0a387bd64222e831fd610270cd7ea2505549758bf75c05a994a6d034f65f8f0e6fdcaeab1a34d4a6b4b636e070a38bce737"
            Test.Use HexString(&HAA, 131), "Test Using Larger Than Block-Size Key - Hash Key First", "80b24263c7c1a3ebb71493c1dd7be8b49b46d1f41b4aeec1121b013783f8f3526b56d037e05f2598bd0fd2215d6a1e5295e64f73f63f0aec8b915a985d786598"
            Test.Use HexString(&HAA, 131), "This is a test using a larger than block-size key and a larger than block-size data. The key needs to be hashed before being used by the HMAC algorithm.", "e37b6a775dc87dbaa4dfa9f96e5e3ffddebd71f8867289865df5a32d20cdc944b6022cac3c4982b10d5eeb55c3e4de15134676fb6de0446065c97440fa8c6a58"
            
        Case "ComputeHash_WithHexValues_ReturnsExpected"
            Test.Use "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", HexString(&HDD, 50), "fa73b0089d56a284efb0f0756c890be9b1b5dbdd8ee81a3655f83e33b2279d39bf3e848279a722c806b485a47e67c807b946a337bee8942674278859e13292fb"
            Test.Use "0102030405060708090a0b0c0d0e0f10111213141516171819", HexString(&HCD, 50), "b0ba465637458c6990e5a8c5f61d4af7e576d97ff94b872de76f8050361ee3dba91ca5c11aa25eb4d679275cc5788063a5f19741120c4f2de2adebeb10a298dd"
                        
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
    Assert.That CorArray.Length(Actual), Iz.EqualTo(128)
    Assert.That Actual, Has.Some.Not.EqualTo(0)
End Sub

Public Sub Constructor_WithNullKey_ThrowsArgumentNullException()
    Dim NullKey() As Byte
    
    On Error Resume Next
    Cor.NewHMACSHA512 NullKey
    
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
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("87aa7cdea5ef619d4ff0b4241a1d6cb02379f4e2ce4ec2787ad0b30545e17cdedaa833b7d6b8a702038b274eaea3f4e4be9d914eeb61f1702e696c203a126854"))
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
        Set CreateHMAC = Cor.NewHMACSHA512(HexToBytes(Key))
    Else
        Set CreateHMAC = New HMACSHA512
    End If
End Function




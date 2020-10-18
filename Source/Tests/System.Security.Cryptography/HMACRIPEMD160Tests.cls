VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACRIPEMD160Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' HMACRIPEMD160Tests
'
' Test vector source: http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
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
            Test.Use "00112233445566778899aabbccddeeff01234567", "", "cf387677bfda8483e63b57e06c3b5ecd8b7fc055"
            Test.Use "00112233445566778899aabbccddeeff01234567", "a", "0d351d71b78e36dbb7391c810a0d2b6240ddbafc"
            Test.Use "00112233445566778899aabbccddeeff01234567", "abc", "f7ef288cb1bbcc6160d76507e0a3bbf712fb67d6"
            Test.Use "00112233445566778899aabbccddeeff01234567", "message digest", "f83662cc8d339c227e600fcd636c57d2571b1c34"
            Test.Use "00112233445566778899aabbccddeeff01234567", "abcdefghijklmnopqrstuvwxyz", "843d1c4eb880ac8ac0c9c95696507957d0155ddb"
            Test.Use "00112233445566778899aabbccddeeff01234567", "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "60f5ef198a2dd5745545c1f0c47aa3fb5776f881"
            Test.Use "00112233445566778899aabbccddeeff01234567", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", "e49c136a9e5627e0681b808a3b97e6a6e661ae79"
            Test.Use "00112233445566778899aabbccddeeff01234567", RepeatString("1234567890", 8), "31be3cc98cee37b79b0619e3e1c2be4f1aa56e6c"
#If Release Then
            Test.Use "00112233445566778899aabbccddeeff01234567", String$(1000000, "a"), "c2aa88c6405658dc225e485488371fb2433fa735"
#End If
            
            Test.Use "0123456789abcdeffedcba987654321000112233", "", "fe69a66c7423eea9c8fa2eff8d9dafb4f17a62f5"
            Test.Use "0123456789abcdeffedcba987654321000112233", "a", "85743e899bc82dbfa36faaa7a25b7cfd372432cd"
            Test.Use "0123456789abcdeffedcba987654321000112233", "abc", "6e4afd501fa6b4a1823ca3b10bd9aa0ba97ba182"
            Test.Use "0123456789abcdeffedcba987654321000112233", "message digest", "2e066e624badb76a184c8f90fba053330e650e92"
            Test.Use "0123456789abcdeffedcba987654321000112233", "abcdefghijklmnopqrstuvwxyz", "07e942aa4e3cd7c04dedc1d46e2e8cc4c741b3d9"
            Test.Use "0123456789abcdeffedcba987654321000112233", "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq", "b6582318ddcfb67a53a67d676b8ad869aded629a"
            Test.Use "0123456789abcdeffedcba987654321000112233", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", "f1be3ee877703140d34f97ea1ab3a07c141333e2"
            Test.Use "0123456789abcdeffedcba987654321000112233", RepeatString("1234567890", 8), "85f164703e61a63131be7e45958e0794123904f9"
#If Release Then
            Test.Use "0123456789abcdeffedcba987654321000112233", String$(1000000, "a"), "82a504a002ba6e6c67f3cd67cedb66dc169bab7a"
#End If
                        
    End Select
End Sub

Public Sub ComputeHash_WithTextValues_ReturnsExpected(ByVal Key As String, ByVal Message As String, ByVal Expected As String)
    Dim csp As HMAC
    Set csp = CreateHMAC(Key)
    
    Dim Actual() As Byte
    Actual = csp.ComputeHash(Encoding.UTF8.GetBytes(Message))
    
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
    Cor.NewHMACRIPEMD160 NullKey
    
    AssertArgumentNullException Err, "Key"
End Sub

Public Sub TransformBlock_WhenCalled_HashesValue()
    Dim csp As HMAC
    Dim Block1() As Byte
    Dim Block2() As Byte
    Set csp = CreateHMAC("00112233445566778899aabbccddeeff01234567")
    Block1 = Encoding.UTF8.GetBytes("message ")
    Block2 = Encoding.UTF8.GetBytes("digest")
    
    Dim Actual() As Byte
    csp.TransformBlock Block1, 0, 8, Block1, 0
    csp.TransformFinalBlock Block2, 0, 6
    Actual = csp.Hash
    
    Assert.That Actual, Iz.EqualTo(HexToBytes("f83662cc8d339c227e600fcd636c57d2571b1c34"))
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
        Set CreateHMAC = Cor.NewHMACRIPEMD160(HexToBytes(Key))
    Else
        Set CreateHMAC = New HMACRIPEMD160
    End If
End Function


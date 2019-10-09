VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' GuidTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithFormats_ReturnsExpected"
            Test.Use "", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "B", "{c200e360-38c5-11ce-ae62-08002b2b79ef}"
            Test.Use "b", "{c200e360-38c5-11ce-ae62-08002b2b79ef}"
            Test.Use "P", "(c200e360-38c5-11ce-ae62-08002b2b79ef)"
            Test.Use "p", "(c200e360-38c5-11ce-ae62-08002b2b79ef)"
            Test.Use "D", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "d", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "N", "c200e36038c511ceae6208002b2b79ef"
            Test.Use "n", "c200e36038c511ceae6208002b2b79ef"
            Test.Use "X", "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}"
            Test.Use "x", "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}"
            
        Case "Parse_WithValue_ReturnsExpected"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use " C200E36038C511CEAE6208002B2B79EF ", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            
        Case "TryParseExact_WithValidValue_ReturnsGuid"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "b"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "B"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "p"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "P"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "n"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "N"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "d"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "D"
            Test.Use "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}", "x"
            Test.Use "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}", "X"
            Test.Use " { 0xc20 0e360 , 0x38c 5,0x11ce, {0xae, 0x62,0x08,0x00,0x2b,0x2b,0x79,0xef} } ", "X"
            
        Case "TryParseExact_WithInvalidValue_ReturnsEmptyGuid"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "p"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "b"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "d"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "n"
            Test.Use "{c200e360 38c5 11ce ae62 08002b2b79ef}", "b"
            Test.Use "(C200E360 38C5 11CE AE62 08002B2B79EF)", "p"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF", "P"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef", "B"
            Test.Use "C200E36038C511CEAE6208002B2B79E", "N"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79E", "d"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", ""
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "z"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79Z", "d"
        
        Case "TryParse_WithInvalidValue_ReturnsFalse"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79eK}"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EFF"
            Test.Use "C200E36038C511CEAE6208002B2B79 F"
            Test.Use "C200E360-38C5-11CE-AE62 08002B2B79EF"
            Test.Use "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}"
            Test.Use "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,Xxef}}"
        
        Case "TryParse_WithValidValue_ReturnsGuid"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)"
            Test.Use "C200E36038C511CEAE6208002B2B79EF"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF"
            Test.Use "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}"
            
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EE}", False
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EF}", True
            
        Case "Equals_WithNonGuidValues_ReturnsFalse"
            Test.Use ""
            Test.Use "abc"
            Test.Use Nothing
            Test.Use Me
            
        Case "CompareTo_WithValues_ReturnsExpected"
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79EE}").Expect 0
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79EF}").Expect Iz.Negative
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79ED}").Expect Iz.Positive
            
        Case "CompareTo_WithAnyValueAgainstEmpty_ReturnsPositive"
            Test.Use "{F0000000-0000-0000-0000-000000000000}"
            Test.Use "{00000000-F000-0000-0000-000000000000}"
            Test.Use "{00000000-0000-F000-0000-000000000000}"
            Test.Use "{00000000-0000-0000-F000-000000000000}"

        Case "GetHashCode_WhenCalled_ReturnsExpected"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79EE}", True
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79Ef}", False
            
        Case "NewGuid_WithValue_ReturnsExpected"
            Test.Use NewBytes(&H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF), "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EF}", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use " C200E36038C511CEAE6208002B2B79EF", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            Test.Use " { 0xc20 0e360 , 0x38c 5,0x11ce, {0xae, 0x62,0x08,0x00,0x2b,0x2b,0x79,0xef} } ", "c200e360-38c5-11ce-ae62-08002b2b79ef"
            
        Case "ToString_WithInvalidFormatSpecifier_ThrowsFormatException"
            Test.Use "f"
            Test.Use "BB"
            
    End Select
End Sub

Public Sub ToString_WithFormats_ReturnsExpected(ByVal GuidFormat As String, ByVal Expected As String)
    Dim GuidToUse As Guid
    Set GuidToUse = Guid.Parse("{c200e360-38c5-11ce-ae62-08002b2b79ef}")
    
    Dim Actual As String
    Actual = GuidToUse.ToString(GuidFormat)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Parse_WithValue_ReturnsExpected(ByVal s As String, ByVal Expected As String)
    Dim Actual As CorLib.Guid
    
    Set Actual = Guid.Parse(s)
    
    Assert.That Actual.ToString, Iz.EqualTo(Expected)
End Sub

Public Sub Parse_WithInvalidValue_ThrowsFormatException()
    On Error Resume Next
    
    Guid.Parse "[c200e360-38c5-11ce-ae62-08002b2b79ef]"
    
    AssertFormatException Err
End Sub

Public Sub ParseExact_WithInvalidFormatSpecifier_ThrowsFormatException()
    On Error Resume Next
    
    Guid.ParseExact "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "z"
    
    AssertFormatException Err
End Sub

Public Sub ParseExact_WithInvalidValue_ThrowsFormatException()
    On Error Resume Next
    
    Guid.ParseExact "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "P"
    
    AssertFormatException Err
End Sub

Public Sub ParseExact_WithWrongLengthValue_ThrowFormatException()
    On Error Resume Next
    
    Guid.ParseExact "{c200e360-38c5-11ce-ae62-08002b2b79ef", "b"
    
    AssertFormatException Err
End Sub

Public Sub TryParseExact_WithValidValue_ReturnsGuid(ByVal s As String, ByVal Format As String)
    Dim Actual As Guid
    Dim Result As Boolean
    
    Result = Guid.TryParseExact(s, Format, Actual)
    
    Dim Expected As Guid
    Set Expected = Guid.Parse("{c200e360-38c5-11ce-ae62-08002b2b79ef}")
    Assert.That Result, Iz.True
    Assert.That Actual, Equals(Expected)
End Sub

Public Sub TryParseExact_WithInvalidValue_ReturnsEmptyGuid(ByVal s As String, ByVal Format As String)
    Dim Actual As Guid
    Dim Result As Boolean
    
    Result = Guid.TryParseExact(s, Format, Actual)
    
    Assert.That Result, Iz.False
    Assert.That Actual, Equals(Guid.EmptyGuid)
End Sub

Public Sub TryParse_WithValidValue_ReturnsGuid(ByVal s As String)
    Dim Actual As Guid
    Dim Result As Boolean
    
    Result = Guid.TryParse(s, Actual)
    
    Const Expected As String = "{c200e360-38c5-11ce-ae62-08002b2b79ef}"
    Assert.That Result, Iz.True
    Assert.That Actual.ToString("b"), Iz.EqualTo(Expected)
End Sub

Public Sub TryParse_WithInvalidValue_ReturnsFalse(ByRef InvalidInput As String)
    Dim Actual As Boolean
    Dim Result As Guid
    
    Actual = Guid.TryParse(InvalidInput, Result)
    
    Assert.That Actual, Iz.False
    Assert.That Result, Equals(Guid.EmptyGuid)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByRef s As String, ByVal Expected As Boolean)
    Dim g As Guid
    Set g = Cor.NewGuid("{C200E360-38C5-11CE-AE62-08002B2B79EF}")
    Dim Other As Guid
    Set Other = Cor.NewGuid(s)
    
    Dim Actual As Boolean
    Actual = g.Equals(Other)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithNonGuidValues_ReturnsFalse(ByVal Other As Variant)
    Dim GuidToCompareTo As New Guid
    
    Dim Actual As Boolean
    Actual = GuidToCompareTo.Equals(Other)
    
    Assert.That Actual, Iz.False
End Sub

Public Function CompareTo_WithValues_ReturnsExpected(ByVal s As String) As Long
    Dim g As Guid
    Set g = Cor.NewGuid("{C200E360-38C5-11CE-AE62-08002B2B79EE}")
    Dim Other As Guid
    Set Other = Cor.NewGuid(s)
    
    Dim Actual As Long
    Actual = g.CompareTo(Other)
    
    CompareTo_WithValues_ReturnsExpected = Actual
End Function

Public Sub CompareTo_WithAnyValueAgainstEmpty_ReturnsPositive(ByVal s As String)
    Dim g As Guid
    Set g = NewGuid(s)
    
    Dim Actual As Long
    Actual = g.CompareTo(Guid.EmptyGuid)
    
    Assert.That Actual, Iz.Positive
End Sub

Public Sub GetHashCode_WhenCalled_ReturnsExpected(ByVal Guid1 As String, ByVal Guid2 As String, ByVal Expected As Boolean)
    Dim Guid1Guid As Guid
    Dim Guid2Guid As Guid
    Set Guid1Guid = Guid.Parse(Guid1)
    Set Guid2Guid = Guid.Parse(Guid2)
    
    Dim Actual As Boolean
    Actual = (Guid1Guid.GetHashCode = Guid2Guid.GetHashCode)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToByteArray_WhenCalled_ReturnsExpected()
    Dim g As Guid
    Set g = Guid.Parse("{C200E360-38C5-11CE-AE62-08002B2B79EF}")
    
    Dim Actual() As Byte
    Actual = g.ToByteArray
    
    Dim Expected() As Byte
    Expected = NewBytes(&H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewGuid_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Actual As Guid
    
    Set Actual = Cor.NewGuid(Value)
    
    Assert.That Actual.ToString, Iz.EqualTo(Expected)
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub NewGuid_WithMemoryPointer_ReturnsGuid()
    Dim Expected() As Byte
    Expected = NewBytes(&H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF)
    
    Dim Actual As Guid
    Set Actual = Cor.NewGuid(VarPtr(Expected(0)))
    
    Assert.That Actual.ToByteArray, Iz.EqualTo(Expected)
End Sub

Public Sub NewGuid_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Byte
    
    On Error Resume Next
    NewGuid NullArray
    
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub ToString_WithInvalidFormatSpecifier_ThrowsFormatException(ByVal InvalidFormat As String)
    On Error Resume Next
    
    Guid.NewGuid.ToString InvalidFormat
    
    AssertFormatException Err
End Sub

Public Sub Serialization_WhenCalled_SerializesInstance()
    Dim Original As CorLib.Guid
    Set Original = Guid.NewGuid
    
    Dim Actual As Guid
    Set Actual = Persist(Original)
    
    Assert.That Actual, Equals(Original)
    Assert.That Actual.IsReadOnly, Iz.True
End Sub

Public Sub EmptyGuid_WhenCalled_ReturnsGuidWithZeroValue()
    Dim Actual As Guid
    
    Set Actual = Guid.EmptyGuid
    
    Dim Expected As Guid
    Set Expected = NewGuid("00000000-0000-0000-0000-000000000000")
    Assert.That Actual, Equals(Expected)
End Sub



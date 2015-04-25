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

Private Const TEST_GUID As String = "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
Private Const TEST_GUID_LOWER As String = "{C200E360-38C5-11CE-AE62-08002B2B79EE}"
Private Const TEST_GUID_UPPER As String = "{C200E360-38C5-11CE-AE62-08002B2B79FF}"

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToString_WithFormats_ReturnsExpected"
            Test.Use "", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "B", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "P", "(C200E360-38C5-11CE-AE62-08002B2B79EF)"
            Test.Use "D", "C200E360-38C5-11CE-AE62-08002B2B79EF"
            Test.Use "N", "C200E36038C511CEAE6208002B2B79EF"
            Test.Use "X", "{0xc200e360,0x38c5,0x11ce,{0xae,0x62,0x08,0x00,0x2b,0x2b,0x79,0xef}}"
            
        Case "Parse_WithValues_ReturnsExpected"
            Test.Use "{c200e360-38c5-11ce-ae62-08002b2b79ef}", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "C200E36038C511CEAE6208002B2B79EF", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "(C200E360-38C5-11CE-AE62-08002B2B79EF)", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "C200E360-38C5-11CE-AE62-08002B2B79EF", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            
        Case "Equals_WithValues_ReturnsExpected"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EF}", "{C200E360-38C5-11CE-AE62-08002B2B79EE}", False
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EF}", "{C200E360-38C5-11CE-AE62-08002B2B79EF}", True
            
        Case "Equals_WithNonGuidValues_ReturnsFalse"
            Test.Use ""
            Test.Use "abc"
            Test.Use Nothing
            Test.Use Me
            
        Case "CompareTo_WithValues_ReturnsExpected"
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79EE}").Expect 0
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79EF}").Expect Iz.Negative
            Test.Use("{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79ED}").Expect Iz.Positive
            
        Case "GetHashCode_WhenCalled_ReturnsExpected"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79EE}", True
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EE}", "{C200E360-38C5-11CE-AE62-08002B2B79Ef}", False
            
        Case "NewGuid_WithValue_ReturnsExpected"
            Test.Use NewBytes(&H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF), "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            Test.Use "{C200E360-38C5-11CE-AE62-08002B2B79EF}", "{C200E360-38C5-11CE-AE62-08002B2B79EF}"
            
    End Select
End Sub

Public Sub ToString_WithFormats_ReturnsExpected(ByVal GuidFormat As String, ByVal Expected As String)
    Dim GuidToUse As Guid
    Set GuidToUse = Guid.Parse(TEST_GUID)
    
    Dim Actual As String
    Actual = GuidToUse.ToString(GuidFormat)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Parse_WithValues_ReturnsExpected(ByVal s As String, ByVal Expected As String)
    Dim Actual As VBCorLib.Guid
    
    Set Actual = Guid.Parse(s)
    
    Assert.That Actual.ToString, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithValues_ReturnsExpected(ByVal Left As String, ByVal Right As String, ByVal Expected As Boolean)
    Dim LeftGuid    As Guid
    Dim RightGuid   As Guid
    Set LeftGuid = Guid.Parse(Left)
    Set RightGuid = Guid.Parse(Right)
    
    Dim Actual As Boolean
    Actual = LeftGuid.Equals(RightGuid)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub Equals_WithNonGuidValues_ReturnsFalse(ByVal Other As Variant)
    Dim GuidToCompareTo As New Guid
    
    Dim Actual As Boolean
    Actual = GuidToCompareTo.Equals(Other)
    
    Assert.That Actual, Iz.False
End Sub

Public Function CompareTo_WithValues_ReturnsExpected(ByVal Left As String, ByVal Right As String) As Long
    Dim LeftGuid As Guid
    Dim RightGuid As Guid
    Set LeftGuid = Guid.Parse(Left)
    Set RightGuid = Guid.Parse(Right)
    
    Dim Actual As Long
    Actual = LeftGuid.CompareTo(RightGuid)
    
    CompareTo_WithValues_ReturnsExpected = Actual
End Function

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
    Dim GuidValue As VBCorLib.Guid
    Set GuidValue = Guid.Parse("{C200E360-38C5-11CE-AE62-08002B2B79EF}")
    
    Dim Actual() As Byte
    Actual = GuidValue.ToByteArray
    
    Dim Expected() As Byte
    Expected = NewBytes(&H60, &HE3, &H0, &HC2, &HC5, &H38, &HCE, &H11, &HAE, &H62, &H8, &H0, &H2B, &H2B, &H79, &HEF)
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NewGuid_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As String)
    Dim Actual As Guid
    
    Set Actual = Cor.NewGuid(Value)
    
    Dim ExpectedGuid As Guid
    Set ExpectedGuid = Guid.Parse(Expected)
    Assert.That Actual, Equals(ExpectedGuid)
End Sub

Public Sub NewGuid_WithNullArray_ThrowsArgumentNullException()
    Dim NullArray() As Byte
    
    On Error GoTo errTrap
    NewGuid NullArray
    
errTrap:
    AssertArgumentNullException Err, "Value"
End Sub

Public Sub TestPersistence()
    Dim G1 As VBCorLib.Guid
    Dim G2 As VBCorLib.Guid
    Dim pp1 As New PropertyBag
    Dim pp2 As New PropertyBag
    
    Set G1 = Guid.NewGuid
    pp1.WriteProperty "g", G1
    pp2.Contents = pp1.Contents
    
    Set G2 = pp2.ReadProperty("g")
    
    Assert.IsTrue G2.Equals(G1), "Guids should be equal."
    
End Sub

Public Sub EmptyGuid_WhenCalled_ReturnsGuidWithZeroValue()
    Assert.AreEqual "{00000000-0000-0000-0000-000000000000}", Guid.EmptyGuid.ToString
End Sub



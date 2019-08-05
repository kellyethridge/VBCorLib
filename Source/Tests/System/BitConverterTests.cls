VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitConverterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' BitConverterTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithValue_ReturnsExpected"
            Test.Use True, NewBytes(1)
            Test.Use False, NewBytes(0)
            Test.Use 1, NewBytes(1, 0)
            Test.Use -1, NewBytes(255, 255)
            Test.Use &H8000, NewBytes(&H0, &H80)
            Test.Use 1&, NewBytes(1, 0, 0, 0)
            Test.Use -1&, NewBytes(255, 255, 255, 255)
            Test.Use &H80000000, NewBytes(&H0, &H0, &H0, &H80)
            Test.Use 1#, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &HF0, &H3F)
            Test.Use -1#, NewBytes(&H0, &H0, &H0, &H0, &H0, &H0, &HF0, &HBF)
            Test.Use 1.2345E+104, NewBytes(&H38, &H16, &HCA, &HF4, &H16, &H8F, &H8B, &H55)
            Test.Use 1!, NewBytes(&H0, &H0, &H80, &H3F)
            Test.Use -1!, NewBytes(&H0, &H0, &H80, &HBF)
            Test.Use 1.2345E+10!, NewBytes(&H70, &HF4, &H37, &H50)
            Test.Use 0.0001@, NewBytes(1, 0, 0, 0, 0, 0, 0, 0)
            Test.Use -922337203685477.5807@, NewBytes(&H1, &H0, &H0, &H0, &H0, &H0, &H0, &H80)
            Test.Use 922337203685477.5807@, NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H7F)
            Test.Use #1/2/2001 8:30:45 AM#, NewBytes(&H9A, &H99, &H99, &H59, &HAB, &H3, &HE2, &H40)
            Test.Use CDec("12345.678911"), NewBytes(&H3F, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H0)
            Test.Use CDec("-12345.678911"), NewBytes(&H3F, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H80)
            Test.Use Int64.One, NewBytes(1, 0, 0, 0, 0, 0, 0, 0)
            Test.Use Int64.MinusOne, NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF)
            
        Case "ToBoolean_WithValue_ReturnsExpected"
            Test.Use NewBytes(0), 0, False
            Test.Use NewBytes(1), 0, True
            Test.Use NewBytes(2), 0, True
            Test.Use NewBytes(0, 1), 1, True
            
        Case "ToBoolean_WithInvalidStartIndex_ThrowsArgumentOutOfRangeException"
            Test.Use 0
            Test.Use 2
    End Select
End Sub

Public Sub GetBytes_WithValue_ReturnsExpected(ByVal Value As Variant, ByVal Expected As Variant)
    Dim Actual() As Byte
    
    Actual = BitConverter.GetBytes(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToBoolean_WithValue_ReturnsExpected(ByRef Value As Variant, ByVal StartIndex As Long, ByVal Expected As Boolean)
    Dim Bytes() As Byte
    Bytes = Value
    
    Dim Actual As Boolean
    Actual = BitConverter.ToBoolean(Bytes, StartIndex)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToBoolean_WithInvalidStartIndex_ThrowsArgumentOutOfRangeException(ByVal StartIndex As Long)
    Dim Bytes(1 To 1) As Byte
    
    On Error Resume Next
    BitConverter.ToBoolean Bytes, StartIndex
    
    AssertArgumentOutOfRangeException Err, "StartIndex"
End Sub

Public Sub ToDouble_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&HCD, &HCC, &HCC, &HCC, &HCC, &HDC, &H5E, &H40)
    
    Dim Actual As Double
    Actual = BitConverter.ToDouble(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(123.45)
End Sub

Public Sub ToDate_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&HCD, &HCC, &HCC, &HCC, &HCC, &HDC, &H5E, &H40)

    Dim Actual As Date
    Actual = BitConverter.ToDate(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(#5/2/1900 10:48:00 AM#)
End Sub

Public Sub ToInt16_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H55, &H66)
    
    Dim Actual As Integer
    Actual = BitConverter.ToInt16(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(&H6655)
End Sub

Public Sub ToInt32_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H55, &H66, &H77, &H88)
    
    Dim Actual As Long
    Actual = BitConverter.ToInt32(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(&H88776655)
End Sub

Public Sub ToSingle_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H66, &HE6, &HF6, &H42)
    
    Dim Actual As Single
    Actual = BitConverter.ToSingle(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(123.45!)
End Sub

Public Sub ToCurrency_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H3E, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0)
    
    Dim Actual As Currency
    Actual = BitConverter.ToCurrency(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(1234567.891@)
End Sub

Public Sub ToString_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H1, &HEF, &HCD, &HAB)
    
    Dim Actual As String
    Actual = BitConverter.ToString(Bytes)
    
    Assert.That Actual, Iz.EqualTo("01-EF-CD-AB")
End Sub

Public Sub ToString_WithStartIndex_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H1, &HEF, &HCD, &HAB)
    
    Dim Actual As String
    Actual = BitConverter.ToString(Bytes, 2)
    
    Assert.That Actual, Iz.EqualTo("CD-AB")
End Sub

Public Sub ToString_WithStartIndexAndLength_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H1, &HEF, &HCD, &HAB)
    
    Dim Actual As String
    Actual = BitConverter.ToString(Bytes, 1, 2)
    
    Assert.That Actual, Iz.EqualTo("EF-CD")
End Sub

Public Sub ToString_WithSingleByte_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H55)
    
    Dim Actual As String
    Actual = BitConverter.ToString(Bytes)
    
    Assert.That Actual, Iz.EqualTo("55")
End Sub

Public Sub ToDecimal_WithValue_ReturnsExpected()
    Dim Bytes() As Byte
    Bytes = NewBytes(&H3F, &H1C, &HDC, &HDF, &H2, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H6, &H0)
    
    Dim Actual As Variant
    Actual = BitConverter.ToDecimal(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(CDec("12345.678911"))
End Sub

Public Sub ToInt64_WithValue_ReturnsExpected()
    Dim Actual As Int64
    Dim Bytes() As Byte
    Bytes = NewBytes(&HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &HFF, &H7F)
    
    Actual = BitConverter.ToInt64(Bytes, 0)
    
    Assert.That Actual, Iz.EqualTo(Int64.MaxValue)
End Sub


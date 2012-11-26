VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConvertTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConvertTests
'
Option Explicit
Implements ITestCaseSource

Private Const AllByteValuesAsBase64 As String = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="
Private AllByteValuesAsBase64WithLineBreaks As String

Private AllByteValues() As Byte
Private mCulture As CultureInfo

Public Sub FixtureSetup()
    ReDim AllByteValues(0 To 255)
    Dim i As Long
    For i = 0 To 255
        AllByteValues(i) = i
    Next
    
    AllByteValuesAsBase64WithLineBreaks = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4" & vbCrLf & _
                                          "OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3Bx" & vbCrLf & _
                                          "cnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmq" & vbCrLf & _
                                          "q6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj" & vbCrLf & _
                                          "5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToBase64String_WithValues_ReturnsExpected"
            Test.Use "", ""
            Test.Use "f", "Zg=="
            Test.Use "fo", "Zm8="
            Test.Use "foo", "Zm9v"
            Test.Use "foob", "Zm9vYg=="
            Test.Use "fooba", "Zm9vYmE="
            Test.Use "foobar", "Zm9vYmFy"
            
        Case "FromBase64String_WithValues_ReturnsExpected"
            Test.Use "", ""
            Test.Use "Zg==", "f"
            Test.Use "Zm8=", "fo"
            Test.Use "Zm9v", "foo"
            Test.Use "Zm9vYg==", "foob"
            Test.Use "Zm9vYmE=", "fooba"
            Test.Use "Zm9vYmFy", "foobar"
            
        Case "ToString_WithValuesUsingDefaultFormatting_ReturnsExpected"
            Set mCulture = Cor.NewCultureInfo("en-US", False)
            ' String
            Test.Use "", ""
            Test.Use "a", "a"
            
            ' Byte
            Test.Use CByte(0), "0"
            Test.Use CByte(1), "1"
            Test.Use CByte(255), "255"
            
            ' Integer
            Test.Use 0, "0"
            Test.Use 1, "1"
            Test.Use CInt(32767), "32767"
            Test.Use -1, "-1"
            Test.Use CInt(-32768), "-32768"
            
            ' Long
            Test.Use 0&, "0"
            Test.Use 1&, "1"
            Test.Use -1&, "-1"
            Test.Use &H7FFFFFFF, "2147483647"
            Test.Use &H80000000, "-2147483648"
            
            ' Double
            Test.Use 0#, "0"
            Test.Use 1#, "1"
            Test.Use 0.1, "0.1"
            Test.Use 123456789012345#, "123456789012345"
            Test.Use -123456789012345#, "-123456789012345"
            Test.Use 1.79769313486231E+308, "1.79769313486231E+308"
            Test.Use 4.94065645841247E-324, "4.94065645841247E-324"
            Test.Use -1.79769313486231E+308, "-1.79769313486231E+308"
            Test.Use -4.94065645841247E-324, "-4.94065645841247E-324"
            Test.Use MakeImaginary(0), "NaN"
            Test.Use MakeImaginary(1), "Infinity"
            Test.Use MakeImaginary(-1), "-Infinity"
            
            ' Single
            Test.Use 0!, "0"
            Test.Use 0.1!, "0.1"
            Test.Use 1!, "1"
            Test.Use 1234567!, "1234567"
            Test.Use -1234567!, "-1234567"
            Test.Use 1.401298E-45, "1.401298E-45"
            Test.Use 3.402823E+38, "3.402823E+38"
            Test.Use -1.401298E-45, "-1.401298E-45"
            Test.Use -3.402823E+38, "-3.402823E+38"
            
            ' Currency
            Test.Use 0@, "0"
            Test.Use 0.1@, "0.1"
            Test.Use -0.1@, "-0.1"
            Test.Use 922337203685477.5807@, "922337203685477.5807"
            Test.Use -922337203685477.5807@ - 0.0001@, "-922337203685477.5808"
            
            ' Decimal
            Test.Use CDec(0), "0"
            Test.Use CDec(0.1), "0.1"
            Test.Use CDec("79228162514264337593543950335"), "79228162514264337593543950335"
            Test.Use CDec("-79228162514264337593543950335"), "-79228162514264337593543950335"
            Test.Use CDec("7.9228162514264337593543950335"), "7.9228162514264337593543950335"
            Test.Use CDec("-7.9228162514264337593543950335"), "-7.9228162514264337593543950335"
            Test.Use CDec("0.0000000000000000000000000001"), "0.0000000000000000000000000001"
            Test.Use CDec("-0.0000000000000000000000000001"), "-0.0000000000000000000000000001"
            
            ' Date
            Test.Use #12:00:00 AM#, "12/30/1899 12:00:00 AM"
            Test.Use #1/1/100#, "1/1/0100 12:00:00 AM"
            Test.Use #12/31/9999 11:59:59 PM#, "12/31/9999 11:59:59 PM"
            
            ' Boolean
            Test.Use True, "True"
            Test.Use False, "False"
            
            ' Object
            Dim Formattable As New MockFormattable
            Formattable.Result = "abc"
            Test.Use Formattable, "abc"
            Dim Object As New MockObject
            Object.ToStringResult = "123"
            Test.Use Object, "123"
            Test.Use Nothing, ""
            
            ' VBA
            Test.Use Empty, "Empty"
            Test.Use Null, "Null"
            
    End Select
End Sub

Public Sub ToBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual As String
    Dim Bytes() As Byte
    Bytes = StrConv(Source, vbFromUnicode)
    
    Actual = Convert.ToBase64String(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(Source)
    
    Dim Bytes() As Byte
    Bytes = StrConv(Expected, vbFromUnicode)
    Assert.That Actual, Iz.EqualTo(Bytes)
End Sub

Public Sub ToBase64String_WithAllByteValues_ReturnsAllByteValuesAsBase64()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
End Sub

Public Sub ToBase64String_WithAllByteValuesAndInsertLineBreaks_ReturnsAllByteValuesAsBase64WithLineBreaks()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues, InsertLineBreaks:=True)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64WithLineBreaks)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64WithLineBreaks_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64WithLineBreaks)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
End Sub

Public Sub ToBase64String_WithNullArray_Throws()
    Dim Bytes() As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64String_WithIndexLessThanLBound_Throws()
    Dim Bytes(1 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ToBase64String_WithLengthLessThanZero_Throws()
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64String_WithIndexAndLengthGreaterThanUBound_Throws()
    Dim Bytes(0 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToBase64CharArray_WithNullInArray_Throws()
    Dim InArray() As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64CharArray_WithNullOutArray_Throws()
    Dim InArray(0) As Byte
    Dim OutArray() As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "OutArray"
End Sub

Public Sub ToBase64CharArray_WithIndexInLessThanLBound_Throws()
    Dim InArray(1 To 2) As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub ToBase64CharArray_WithIndexOutLessThanLBound_Throws()
    Dim InArray(2) As Byte
    Dim OutArray(1 To 5) As Integer
    
    On Error Resume Next
    
    Convert.ToBase64CharArray InArray, 0, 1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexOut"
End Sub

Public Sub ToBase64CharArray_WithLengthLessThanZero_Throws()
    Dim InArray(1) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, -1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64CharArray_WithIndexInAndLengthGreaterThanUBound_Throws()
    Dim InArray(1) As Byte
    Dim OutArray(2) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub ToBase64CharArray_WithIndexOutAndLengthGreaterThanUBound_Throws()
    Dim InArray(2) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexOut"
End Sub

Public Sub FromBase64CharArray_WithNullArray_Throws()
    Dim InArray() As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 0
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub FromBase64CharArray_WithIndexInLessThanLBound_Throws()
    Dim InArray(1 To 2) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub FromBase64CharArray_WithLengthLessThanZero_Throws()
    Dim InArray(0) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub FromBase64CharArray_WithIndexInAndLengthGreaterThanUBound_Throws()
    Dim InArray(2) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 4
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub FromBase64String_WithMalformedString_Throws()
    Const Malformed As String = "Zm9vYg=" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub FromBase64String_WithTooMuchPadding_Throws()
    Const Malformed As String = "Zm9vYg===" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub FromBase64String_WithInvalidCharater_Throws()
    Const Malformed As String = "Zm9vY.==" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub ToString_WithValuesUsingDefaultFormatting_ReturnsExpected(ByRef Value As Variant, ByVal Expected As String)
    Set Thread.CurrentThread.CurrentCulture = mCulture
    Dim Actual As String
    
    Actual = Convert.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFormattable_SuppliesFormatInformation()
    Dim Mock As New MockFormattable
    
    Convert.ToString Mock, "ab"
    
    Dim Actual As String
    Actual = Mock.Format
    Assert.That Actual, Iz.EqualTo("ab")
End Sub

Public Sub ToString_WithFormattable_SuppliesProvider()
    Dim Mock As New MockFormattable
    Dim Provider As New NumberFormatInfo
    
    Convert.ToString Mock, , Provider
    
    Dim Actual As IFormatProvider
    Set Actual = Mock.Provider
    Assert.That Actual, Iz.SameAs(Provider)
End Sub

Public Sub ToString_WithMissingValue_ReturnsEmptyString()
    Dim Actual As String
    
    Actual = Convert.ToString(Missing)
    
    Assert.That Actual, Iz.Empty
End Sub

Private Function MakeImaginary(ByVal Value As Double) As Double
    On Error Resume Next
    MakeImaginary = Value / 0#
    On Error GoTo 0
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromBase64Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' FromBase64Tests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Const AllByteValuesAsBase64 As String = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="
Private Const AllByteValuesAsBase64WithLineBreaks As String = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4" & vbCrLf & _
                                                              "OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3Bx" & vbCrLf & _
                                                              "cnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmq" & vbCrLf & _
                                                              "q6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj" & vbCrLf & _
                                                              "5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="

Private AllByteValues() As Byte

Public Sub FixtureSetup()
    Set Thread.CurrentThread.CurrentCulture = Cor.NewCultureInfo("en-US", False)
    
    ReDim AllByteValues(0 To 255)
    Dim i As Long
    For i = 0 To 255
        AllByteValues(i) = i
    Next
End Sub

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "FromBase64String_WithValues_ReturnsExpected"
            Test.Use "", ""
            Test.Use "Zg==", "f"
            Test.Use "Zm8=", "fo"
            Test.Use "Zm9v", "foo"
            Test.Use "Zm9vYg==", "foob"
            Test.Use "Zm9vYmE=", "fooba"
            Test.Use "Zm9vYmFy", "foobar"
    End Select
End Sub

Public Sub FromBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(Source)
    
    Dim Bytes() As Byte
    Bytes = StrConv(Expected, vbFromUnicode)
    Assert.That Actual, Iz.EqualTo(Bytes)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64WithLineBreaks_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64WithLineBreaks)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
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
    
    AssertArgumentOutOfRangeException Err, "Offset"
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
    
    AssertArgumentOutOfRangeException Err, "Offset"
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

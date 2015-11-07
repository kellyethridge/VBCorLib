VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToBase64Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ToBase64Tests
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

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System"
End Sub

Public Sub FixtureSetup()
    Set Thread.CurrentThread.CurrentCulture = Cor.NewCultureInfo("en-US", False)
    
    ReDim AllByteValues(0 To 255)
    Dim i As Long
    For i = 0 To 255
        AllByteValues(i) = i
    Next
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
            
    End Select
End Sub

Public Sub ToBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual As String
    Dim Bytes() As Byte
    Bytes = StrConv(Source, vbFromUnicode)
    
    Actual = Convert.ToBase64String(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToBase64String_WithAllByteValues_ReturnsAllByteValuesAsBase64()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64)
End Sub

Public Sub ToBase64String_WithAllByteValuesAndInsertLineBreaks_ReturnsAllByteValuesAsBase64WithLineBreaks()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues, InsertLineBreaks:=True)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64WithLineBreaks)
End Sub

Public Sub ToBase64String_WithNullArray_Throws()
    Dim Bytes() As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64String_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim Bytes(1 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ToBase64String_WithLengthLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64String_WithIndexAndLengthGreaterThanUBound_ThrowsArgumentOutOfRangeException()
    Dim Bytes(0 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToBase64CharArray_WithNullInArray_ThrowsArgumentNullException()
    Dim InArray() As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64CharArray_WithNullOutArray_ThrowsArgumentNullException()
    Dim InArray(0) As Byte
    Dim OutArray() As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "OutArray"
End Sub

Public Sub ToBase64CharArray_WithIndexInLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim InArray(1 To 2) As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub ToBase64CharArray_WithIndexOutLessThanLBound_ThrowsArgumentOutOfRangeException()
    Dim InArray(2) As Byte
    Dim OutArray(1 To 5) As Integer
    
    On Error Resume Next
    
    Convert.ToBase64CharArray InArray, 0, 1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexOut"
End Sub

Public Sub ToBase64CharArray_WithLengthLessThanZero_ThrowsArgumentOutOfRangeException()
    Dim InArray(1) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, -1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64CharArray_WithIndexInAndLengthGreaterThanUBound_ThrowsArgumentOutOfRangeException()
    Dim InArray(1) As Byte
    Dim OutArray(2) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToBase64CharArray_WithIndexOutAndLengthGreaterThanUBound_ThrowsArgumentOutOfRangeException()
    Dim InArray(2) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err
End Sub

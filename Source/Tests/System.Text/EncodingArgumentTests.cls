VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncodingArgumentTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' EncodingArgumentTests
'
Option Explicit
Implements ITestCaseSource
Implements ICategorizable

Private Sub ICategorizable_Categorize(ByVal Cat As SimplyVBComp.CategoryCollector)
    Cat.ForFixture "System.Text"
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithNullArray_ThrowsArgumentNullException"
            AddEncodings Test
    
        Case "GetBytes_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetBytes_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetBytes_WithMissingCount_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetBytes_WithMissingIndex_ThrowsArgumentException"
            AddEncodings Test
        
        Case "GetByteCount_WithMissingCount_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetByteCount_WithMissingIndex_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithCharIndexPlusCharCountBeyondChars_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetBytesEx_WithByteIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
        
        Case "GetBytesEx_WithNotEnoughCapacity_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetCharCount_WithNullArray_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetCharCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharCount_WithIndexPlusCountIsNotInValidRange_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharCount_WithMissingIndex_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetCharCount_WithMissingCount_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetString_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetString_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetString_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetString_WithMissingByteIndex_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetString_WithMissingByteCount_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetString_WithNullBytes_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetChars_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetChars_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetChars_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetChars_WithNullBytes_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetChars_WithMissingByteIndex_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetChars_WithMissingByteCount_ThrowsArgumentException"
            AddEncodings Test
            
        Case "GetCharsEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharsEx_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharsEx_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharsEx_WithNullBytes_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetCharsEx_WithNullChars_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetCharsEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharsEx_WithNotEnoughCapacity_ThrowsArgumentException"
            AddEncodings Test
        
        Case "GetBytesEx_WithNullCharsArray_ThrowsArgumentNullException"
            AddEncodings Test
            
        Case "GetBytesEx_WithNullBytesArray_ThrowsArgumentNullException"
            AddEncodings Test
        
        Case "GetMaxByteCount_WithNegativeValue_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetMaxCharCount_WithNegativeValue_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
            
        Case "GetCharsEx_WithCharIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException"
            AddEncodings Test
    End Select
End Sub

Private Sub AddEncodings(ByVal Test As TestCaseCollector)
    Test.Use New ASCIIEncoding
    Test.Use New UTF7Encoding
    Test.Use New UTF8Encoding
    Test.Use New UnicodeEncoding
    Test.Use Encoding.GetEncoding(437)
    Test.Use Encoding.GetEncoding(936)
End Sub


Public Sub GetBytes_WithNullArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim NullChars() As Integer
    
    On Error Resume Next
    Encoding.GetBytes NullChars
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub GetBytes_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetBytes "abc", -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0 To 1) As Integer
    On Error Resume Next
    
    Encoding.GetBytes Chars, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetBytes_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(1 To 2) As Integer
    On Error Resume Next
    
    Encoding.GetBytes Chars, 1, 3
    
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub

Public Sub GetBytes_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetBytes "abc", 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetBytes_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetBytes "abc", , 2
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetByteCount_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetByteCount NewIntegers(1, 2), Count:=2
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetByteCount_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetByteCount NewIntegers(1, 2), 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetByteCount NewIntegers(1, 2), -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error Resume Next
    
    Encoding.GetByteCount NewIntegers(0), 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(1 To 2) As Integer
    
    On Error Resume Next
    Encoding.GetByteCount Chars, 1, 3
    
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub

Public Sub GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(-2 To 0) As Integer
    Dim Bytes(0 To 2) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx Chars, -3, 1, Bytes, 0
    
    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx Chars, 0, -1, Bytes, 0
    
    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub GetBytesEx_WithCharIndexPlusCharCountBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx Chars, 0, 2, Bytes, 0
    
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub


Public Sub GetBytesEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(-1 To 0) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx Chars, 0, 1, Bytes, -2
    
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetBytesEx_WithByteIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(-1 To 0) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx Chars, 0, 1, Bytes, 1
    
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetBytesEx_WithNotEnoughCapacity_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    ReDim Bytes(1 To 3)
    
    On Error Resume Next
    Encoding.GetBytesEx "12345", 0, 5, Bytes, 1
    
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub GetBytesEx_WithNullCharsArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim NullChars() As Integer
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx NullChars, 0, 1, Bytes, 0
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub GetBytesEx_WithNullBytesArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim NullBytes() As Byte
    
    On Error Resume Next
    Encoding.GetBytesEx "abc", 0, 1, NullBytes, 0
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetCharCount_WithNullArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetCharCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes, -1, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetCharCount_WithIndexPlusCountIsNotInValidRange_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes, 5, 2
    
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetCharCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetCharCount_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes, , 1
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetCharCount_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetCharCount Bytes, 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetString_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes, 4, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetString_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes, 5, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetString_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 4) As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes, 0, 6
    
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetString_WithMissingByteIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes, , 1
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetString_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetString_WithMissingByteCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetString Bytes, 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetChars_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes, 4, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetChars_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes, 5, -1
    
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetChars_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes, 4, 3
    
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetChars_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetChars_WithMissingByteIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes, , 1
    
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetChars_WithMissingByteCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Encoding.GetChars Bytes, 0
    
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetCharsEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, -1, 1, Chars, 0
    
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetCharsEx_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, -1, Chars, 0
    
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub GetCharsEx_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 4, 3, Chars, 0
    
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetCharsEx_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    Dim Chars(10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, 0, Chars, 0
    
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetCharsEx_WithNullChars_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes(10) As Byte
    Dim Chars() As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, 1, Chars, 0
    
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub GetCharsEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(10) As Byte
    Dim Chars(10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, 1, Chars, -1
    
    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub GetCharsEx_WithCharIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 10) As Byte
    Dim Chars(0 To 10) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, 1, Chars, 11
    
    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub GetCharsEx_WithNotEnoughCapacity_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    ReDim Bytes(0 To 8)
    Dim Chars(0 To 2) As Integer
    
    On Error Resume Next
    Encoding.GetCharsEx Bytes, 0, 8, Chars, 0
    
    AssertArgumentException Err, "Chars"
End Sub

Public Sub GetMaxByteCount_WithNegativeValue_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error Resume Next
    Encoding.GetMaxByteCount -1
    
    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub GetMaxCharCount_WithNegativeValue_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error Resume Next
    Encoding.GetMaxCharCount -1
    
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub













































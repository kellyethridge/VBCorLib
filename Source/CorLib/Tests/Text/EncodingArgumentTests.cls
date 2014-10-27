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
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: EncodingArgumentTests
'
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "GetBytes_WithNullArray_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
    
        Case "GetBytes_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetBytes_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetBytes_WithMissingCount_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetBytes_WithMissingIndex_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithMissingCount_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetByteCount_WithMissingIndex_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithCharIndexPlusCharCountBeyondChars_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetBytesEx_WithByteIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
        
        Case "GetBytesEx_WithNotEnoughCapacity_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharCount_WithNullArray_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharCount_WithMissingIndex_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharCount_WithMissingCount_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithMissingByteIndex_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithMissingByteCount_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetString_WithNullBytes_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithNullBytes_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithMissingByteIndex_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetChars_WithMissingByteCount_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithNullBytes_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithNullChars_ThrowsArgumentNullException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException"
            Test.Use New ASCIIEncoding
            
        Case "GetCharsEx_WithNotEnoughCapacity_ThrowsArgumentException"
            Test.Use New ASCIIEncoding
            
    End Select
End Sub


Public Sub GetBytes_WithNullArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim NullChars() As Integer
    On Error GoTo errTrap:
    
    Encoding.GetBytes NullChars
    
errTrap:
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub GetBytes_WithIndexLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap:
    
    Encoding.GetBytes "abc", -1, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetBytes_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap:
    
    Encoding.GetBytes "abc", 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetBytes_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap:
    
    Encoding.GetBytes "abc", 1, 3
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub

Public Sub GetBytes_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error GoTo errTrap:
    
    Encoding.GetBytes "abc", 0
    
errTrap:
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetBytes_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error GoTo errTrap:
    
    Encoding.GetBytes "abc", , 2
    
errTrap:
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetByteCount_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error GoTo errTrap
    
    Encoding.GetByteCount "abc", Count:=2
    
errTrap:
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetByteCount_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    On Error GoTo errTrap
    
    Encoding.GetByteCount "abc", 0
    
errTrap:
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetByteCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap
    
    Encoding.GetByteCount "abc", -1, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetByteCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap
    
    Encoding.GetByteCount "abc", 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetByteCount_WithIndexPlusCountExtendBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    On Error GoTo errTrap
    
    Encoding.GetByteCount "abc", 1, 3
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub

Public Sub GetBytesEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(-2 To 0) As Integer
    Dim Bytes(0 To 2) As Byte
    On Error GoTo errTrap
    
    Encoding.GetBytesEx Chars, -3, 1, Bytes, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub GetBytesEx_WithCharCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetBytesEx Chars, 0, -1, Bytes, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "CharCount"
End Sub

Public Sub GetBytesEx_WithCharIndexPlusCharCountBeyondChars_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetBytesEx Chars, 0, 2, Bytes, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Chars"
End Sub


Public Sub GetBytesEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(-1 To 0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetBytesEx Chars, 0, 1, Bytes, -2
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetBytesEx_WithByteIndexGreaterThanUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Chars(0) As Integer
    Dim Bytes(-1 To 0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetBytesEx Chars, 0, 1, Bytes, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetBytesEx_WithNotEnoughCapacity_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    ReDim Bytes(1 To 3)
    On Error GoTo errTrap
    
    Encoding.GetBytesEx "12345", 0, 5, Bytes, 1
    
errTrap:
    AssertArgumentException Err, "Bytes"
End Sub

Public Sub GetCharCount_WithNullArray_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    On Error GoTo errTrap
    
    Encoding.GetCharCount Bytes
    
errTrap:
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetCharCount_WithIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetCharCount Bytes, -1, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub GetCharCount_WithCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetCharCount Bytes, 0, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Count"
End Sub

Public Sub GetCharCount_WithMissingIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetCharCount Bytes, , 1
    
errTrap:
    AssertArgumentException Err, "Index"
End Sub

Public Sub GetCharCount_WithMissingCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetCharCount Bytes, 0
    
errTrap:
    AssertArgumentException Err, "Count"
End Sub

Public Sub GetString_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes, 4, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetString_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes, 5, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub GetString_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 4) As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes, 0, 6
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetString_WithMissingByteIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes, , 1
    
errTrap:
    AssertArgumentException Err, "ByteIndex"
End Sub

Public Sub GetString_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes
    
errTrap:
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetString_WithMissingByteCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetString Bytes, 0
    
errTrap:
    AssertArgumentException Err, "ByteCount"
End Sub

Public Sub GetChars_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetChars Bytes, 4, 1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetChars_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(5 To 10) As Byte
    On Error GoTo errTrap
    
    Encoding.GetChars Bytes, 5, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub GetChars_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    On Error GoTo errTrap
    
    Encoding.GetChars Bytes, 4, 3
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetChars_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    On Error GoTo errTrap:
    
    Encoding.GetChars Bytes
    
errTrap:
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetChars_WithMissingByteIndex_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetChars Bytes, , 1
    
errTrap:
    AssertArgumentException Err, "ByteIndex"
End Sub

Public Sub GetChars_WithMissingByteCount_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes(0) As Byte
    On Error GoTo errTrap
    
    Encoding.GetChars Bytes, 0
    
errTrap:
    AssertArgumentException Err, "ByteCount"
End Sub

Public Sub GetCharsEx_WithByteIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, -1, 1, Chars, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteIndex"
End Sub

Public Sub GetCharsEx_WithByteCountLessThanZero_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 0, -1, Chars, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "ByteCount"
End Sub

Public Sub GetCharsEx_WithByteIndexPlusByteCountBeyondUBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(0 To 5) As Byte
    Dim Chars(10) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 4, 3, Chars, 0
    
errTrap:
    AssertArgumentOutOfRangeException Err, "Bytes"
End Sub

Public Sub GetCharsEx_WithNullBytes_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    Dim Chars(10) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 0, 0, Chars, 0
    
errTrap:
    AssertArgumentNullException Err, "Bytes"
End Sub

Public Sub GetCharsEx_WithNullChars_ThrowsArgumentNullException(ByVal Encoding As Encoding)
    Dim Bytes(10) As Byte
    Dim Chars() As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 0, 1, Chars, 0
    
errTrap:
    AssertArgumentNullException Err, "Chars"
End Sub

Public Sub GetCharEx_WithCharIndexLessThanLBound_ThrowsArgumentOutOfRangeException(ByVal Encoding As Encoding)
    Dim Bytes(10) As Byte
    Dim Chars(10) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 0, 1, Chars, -1
    
errTrap:
    AssertArgumentOutOfRangeException Err, "CharIndex"
End Sub

Public Sub GetCharsEx_WithNotEnoughCapacity_ThrowsArgumentException(ByVal Encoding As Encoding)
    Dim Bytes() As Byte
    ReDim Bytes(0 To 3)
    Dim Chars(0 To 2) As Integer
    On Error GoTo errTrap
    
    Encoding.GetCharsEx Bytes, 0, 4, Chars, 0
    
errTrap:
    AssertArgumentException Err, "Chars"
End Sub

















































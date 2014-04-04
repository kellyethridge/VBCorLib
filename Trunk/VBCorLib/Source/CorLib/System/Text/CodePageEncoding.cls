VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodePageEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: CodePageEncoding
'

''
' An encoding class that uses the underlying encoding support of Windows.
'

Option Explicit
Implements IObject
Implements Encoding


Private mCodePage       As Long
Private mMaxCharSize    As Long
Private mChars()        As Integer
Private mCharsSA        As SafeArray1d



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal CodePage As Long)
    Dim Info As CPINFO
    
    If GetCPInfo(CodePage, Info) = BOOL_FALSE Then _
        Throw Cor.NewNotSupportedException("Either the Code Page is invalid or the Encoding is not installed.")
    
    mCodePage = CodePage
    mMaxCharSize = Info.MaxCharSize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = modStaticClasses.Encoding.EncodingTable
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    If IsObject(Value) Then
        If TypeOf Value Is Encoding Then
            Dim ec As Encoding
            Set ec = Value
            IObject_Equals = (ec.CodePage = mCodePage)
        End If
    End If
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = mCodePage
End Function

Private Function IObject_ToString() As String
    IObject_ToString = "VBCorLib.Encoding"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_BodyName() As String
    Encoding_BodyName = Table.BodyName(mCodePage)
End Property

Private Property Get Encoding_CodePage() As Long
    Encoding_CodePage = mCodePage
End Property

Private Property Get Encoding_EncodingName() As String
    Encoding_EncodingName = Table.EncodingName(mCodePage)
End Property

Private Function Encoding_Equals(Value As Variant) As Boolean
    Encoding_Equals = IObject_Equals(Value)
End Function

Private Function Encoding_GetByteCount(Chars As Variant, Optional Index As Variant, Optional Count As Variant) As Long
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    Result = GetOptionalArrayRange(SAPtr(mChars), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    If mMaxCharSize = 1 Or ElemCount = 0 Then
        Encoding_GetByteCount = ElemCount
    Else
        Encoding_GetByteCount = WideCharToMultiByte(mCodePage, 0, mChars(ElemIndex), ElemCount, vbNullString, 0, vbNullString, 0)
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    Result = GetOptionalArrayRange(SAPtr(mChars), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim Size As Long
    Size = Encoding_GetMaxByteCount(ElemCount)
    
    Dim Ret() As Byte
    If Size > 0 Then
        ReDim Ret(0 To Size - 1)
        
        Dim ByteCount As Long
        ByteCount = WideCharToMultiByte(mCodePage, 0, mChars(ElemIndex), ElemCount, Ret(0), Size, vbNullString, 0)
        If ByteCount <> Size Then ReDim Preserve Ret(0 To ByteCount - 1)
    Else
        Ret = Cor.NewBytes()
    End If
    
    SAPtr(mChars) = 0
    Encoding_GetBytes = Ret
End Function

Private Function Encoding_GetBytesEx(Chars As Variant, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long) As Long
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(mChars), CharIndex, CharCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Chars", CharIndex, "CharIndex", CharCount, "CharCount")
    
    If CharCount > 0 Then
        Dim RequiredSize As Long
        RequiredSize = WideCharToMultiByte(mCodePage, 0, mChars(CharIndex), CharCount, 0, 0, vbNullString, 0)
        
        Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, RequiredSize)
        Select Case Result
            Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Bytes")
            Case Is <> NO_ERROR:                Call ThrowArrayRangeException(Result, "Bytes", ByteIndex, "ByteIndex", 0, vbNullString)
        End Select
        
        Encoding_GetBytesEx = WideCharToMultiByte(mCodePage, 0, mChars(CharIndex), CharCount, Bytes(ByteIndex), cArray.GetLength(Bytes), vbNullString, 0)
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetCharCount(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Byte", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    If ElemCount = 0 Then Exit Function
    Encoding_GetCharCount = MultiByteToWideChar(mCodePage, 0, Bytes(ElemIndex), ElemCount, ByVal 0&, 0)
End Function

Private Function Encoding_GetChars(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Integer()
    Dim ElemIndex As Long
    Dim ElemCount As Long
    
    Dim Result As Long
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim Size As Long
    Size = Encoding_GetMaxCharCount(ElemCount)
    
    Dim Ret() As Integer
    If Size > 0 Then
        ReDim Ret(0 To Size - 1)
        Size = Encoding_GetCharsEx(Bytes, Index, Count, Ret, 0)
        ReDim Preserve Ret(0 To Size - 1)
    Else
        Ret = Cor.NewIntegers()
    End If
    
    SAPtr(mChars) = 0
    Encoding_GetChars = Ret
End Function

Private Function Encoding_GetCharsEx(Bytes() As Byte, ByVal ByteIndex As Long, ByVal ByteCount As Long, Chars() As Integer, ByVal CharIndex As Long) As Long
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(Bytes), ByteIndex, ByteCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Bytes", ByteIndex, "ByteIndex", ByteCount, "ByteCount")
    
    If ByteCount > 0 Then
        Dim RequiredSize As Long
        RequiredSize = MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, ByVal 0&, 0)
        
        Result = VerifyArrayRange(SAPtr(Chars), CharIndex, RequiredSize)
        Select Case Result
            Case Argument_InvalidOffLen:   Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "Chars")
            Case Is <> NO_ERROR:                Call ThrowArrayRangeException(Result, "Chars", CharIndex, "CharIndex", 0, vbNullString)
        End Select
        
        Encoding_GetCharsEx = MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, Chars(CharIndex), cArray.GetLength(Chars))
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetDecoder() As Decoder
    Dim Ret As New CodePageDecoder
    Call Ret.Init(mMaxCharSize, mCodePage)
    Set Encoding_GetDecoder = Ret
End Function

Private Function Encoding_GetEncoder() As Encoder
    Dim Ret As New DefaultEncoder
    Call Ret.Init(Me)
    Set Encoding_GetEncoder = Ret
End Function

Private Function Encoding_GetHashCode() As Long
    Encoding_GetHashCode = mCodePage
End Function

Private Function Encoding_GetMaxByteCount(ByVal CharCount As Long) As Long
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("CharCount", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
        
    Encoding_GetMaxByteCount = CharCount * mMaxCharSize
End Function

Private Function Encoding_GetMaxCharCount(ByVal ByteCount As Long) As Long
    If ByteCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("ByteCount", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    Encoding_GetMaxCharCount = ByteCount
End Function

Private Function Encoding_GetPreamble() As Byte()
    Encoding_GetPreamble = Cor.NewBytes()
End Function

Private Function Encoding_GetString(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As String
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index))
    
    Dim Size As Long
    Size = Encoding_GetMaxCharCount(Count)
    
    Dim Ret As String
    Ret = SysAllocStringLen(0, Size)
    Call AttachChars(Ret, mChars, mCharsSA)
    Size = Encoding_GetCharsEx(Bytes, ElemIndex, ElemCount, mChars, 0)
    SAPtr(mChars) = 0
    Encoding_GetString = Left$(Ret, Size)
End Function

Private Property Get Encoding_HeaderName() As String
    Encoding_HeaderName = Table.HeaderName(mCodePage)
End Property

Private Property Get Encoding_IsBrowserDisplay() As Boolean
    Encoding_IsBrowserDisplay = Table.IsBrowserDisplay(mCodePage)
End Property

Private Property Get Encoding_IsBrowserSave() As Boolean
    Encoding_IsBrowserSave = Table.IsBrowserSave(mCodePage)
End Property

Private Property Get Encoding_IsMailNewsDisplay() As Boolean
    Encoding_IsMailNewsDisplay = Table.IsMailNewsDisplay(mCodePage)
End Property

Private Property Get Encoding_IsMailNewsSave() As Boolean
    Encoding_IsMailNewsSave = Table.IsMailNewsSave(mCodePage)
End Property

Private Function Encoding_ToString() As String
    Encoding_ToString = "VBCorLib.Encoding"
End Function

Private Property Get Encoding_WebName() As String
    Encoding_WebName = Table.WebName(mCodePage)
End Property

Private Property Get Encoding_WindowsCodePage() As Long
    Encoding_WindowsCodePage = Table.WindowsCodePage(mCodePage)
End Property

Private Property Get Encoding_IsSingleByte() As Boolean
    Encoding_IsSingleByte = False
End Property



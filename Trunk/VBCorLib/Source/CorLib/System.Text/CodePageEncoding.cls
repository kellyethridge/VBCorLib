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

Private mIsReadOnly         As Boolean
Private mCodePage           As Long
Private mMaxCharSize        As Long
Private mDecoderFallback    As DecoderFallback
Private mEncoderFallback    As EncoderFallback
Private mChars()            As Integer
Private mCharsSA            As SafeArray1d


Friend Property Get CodePage() As Long
    CodePage = mCodePage
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal CodePage As Long)
    Dim Info As CPINFO
    
    If GetCPInfo(CodePage, Info) = BOOL_FALSE Then _
        Throw Cor.NewNotSupportedException("Either the Code Page is invalid or the Encoding is not installed.")
    
    mCodePage = CodePage
    mMaxCharSize = Info.MaxCharSize
End Sub

Friend Sub InitClone(ByVal Other As CodePageEncoding)

End Sub

Private Sub Class_Terminate()
    SAPtr(mChars) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Table() As EncodingTable
    Set Table = Statics.Encoding.EncodingTable
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    Dim Other As CodePageEncoding
    
    On Error GoTo Catch
    Set Other = Value
    IObject_Equals = Other.CodePage = mCodePage
    
Catch:
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = mCodePage
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoding Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get Encoding_IsReadOnly() As Boolean
    Encoding_IsReadOnly = mIsReadOnly
End Property

Private Function Encoding_Clone() As Encoding
    Dim Clone As New CodePageEncoding
    Clone.InitClone Me
    Set Encoding_Clone = Clone
End Function

Private Property Set Encoding_DecoderFallback(ByVal Value As DecoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mDecoderFallback = Value
End Property

Private Property Get Encoding_DecoderFallback() As DecoderFallback
    Set Encoding_DecoderFallback = mDecoderFallback
End Property

Private Property Set Encoding_EncoderFallback(ByVal Value As EncoderFallback)
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    
    Set mEncoderFallback = Value
End Property

Private Property Get Encoding_EncoderFallback() As EncoderFallback
    Set Encoding_EncoderFallback = mEncoderFallback
End Property

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
    Dim LocalChars() As Integer
    Dim Range           As ListRange
    
    LocalChars = AsChars(Chars)
    
    On Error GoTo Catch
    Range = ValidateGetBytes(LocalChars, Index, Count)
    
    If mMaxCharSize = 1 Or Range.Count = 0 Then
        Encoding_GetByteCount = Range.Count
    Else
        Encoding_GetByteCount = WideCharToMultiByte(mCodePage, 0, LocalChars(Range.Index), Range.Count, vbNullString, 0, vbNullString, 0)
    End If
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars LocalChars
End Function

Private Function Encoding_GetBytes(Chars As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    Call AttachChars(Chars, mChars, mCharsSA)
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Range As ListRange
    
    Range = MakeArrayRange(mChars, Index, Count)
    ElemIndex = Range.Index
    ElemCount = Range.Count
    
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
    
    ValidateArrayRange mChars, CharIndex, CharCount, NameOfChars, NameOfCharIndex, NameOfCharCount
    
    
    If CharCount > 0 Then
        Dim RequiredSize As Long
        RequiredSize = WideCharToMultiByte(mCodePage, 0, mChars(CharIndex), CharCount, 0, 0, vbNullString, 0)
        ValidateArrayRange Bytes, ByteIndex, RequiredSize
                
        Encoding_GetBytesEx = WideCharToMultiByte(mCodePage, 0, mChars(CharIndex), CharCount, Bytes(ByteIndex), CorArray.Length(Bytes), vbNullString, 0)
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetCharCount(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Long
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Range As ListRange
    
    Range = MakeArrayRange(Bytes, Index, Count)
    ElemIndex = Range.Index
    ElemCount = Range.Count
    
    If ElemCount = 0 Then Exit Function
    Encoding_GetCharCount = MultiByteToWideChar(mCodePage, 0, Bytes(ElemIndex), ElemCount, ByVal 0&, 0)
End Function

Private Function Encoding_GetChars(Bytes() As Byte, Optional Index As Variant, Optional Count As Variant) As Integer()
    Dim ElemIndex As Long
    Dim ElemCount As Long
    Dim Range As ListRange
    
    Range = MakeArrayRange(Bytes, Index, Count)
    ElemIndex = Range.Index
    ElemCount = Range.Count
        
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
    ValidateByteRange Bytes, ByteIndex, ByteCount, NameOfBytes, NameOfByteIndex, NameOfByteCount
    
    If ByteCount > 0 Then
        Dim RequiredSize As Long
        RequiredSize = MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, ByVal 0&, 0)
        ValidateCharRange Chars, CharIndex, RequiredSize
        
        Encoding_GetCharsEx = MultiByteToWideChar(mCodePage, 0, Bytes(ByteIndex), ByteCount, Chars(CharIndex), CorArray.Length(Chars))
    End If
    SAPtr(mChars) = 0
End Function

Private Function Encoding_GetDecoder() As Decoder
    Dim Ret As New CodePageDecoder
    Call Ret.Init(mMaxCharSize, mCodePage)
    Set Encoding_GetDecoder = Ret
End Function

Private Function Encoding_GetEncoder() As Encoder
    Set Encoding_GetEncoder = Cor.NewDefaultEncoder(Me)
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
    Dim Range As ListRange
    
    Range = MakeArrayRange(Bytes, Index, Count)
    ElemIndex = Range.Index
    ElemCount = Range.Count
    
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



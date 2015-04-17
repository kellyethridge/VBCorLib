VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA1Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: SHA1Managed
'

''
' Provides services to compute the SHA-1 hash value.
'
' @remarks The SHA-1 algorithm is implemented in Visual Basic.
' <p>Details can be found at http://csrc.nist.gov/CryptoToolkit/tkhash.html</p>
' @see SHA1CryptoServiceProvider
' @see SHA1
' @see SHA1Static
' @see HashAlgorithm
'
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements SHA1
Implements ICryptoTransform
Implements IHashAlgorithm

Private mBase               As HashAlgorithmBase
Private mW(0 To 79)         As Long
Private mState(0 To 4)      As Long



''
' Returns if this service provider can be reused to transfrom more data.
'
' @return Always returns True.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if this transform can operate on multiple blocks at once.
'
' @return Always returns True.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = True
End Property

''
' Returns the computed hash.
'
' @return A 20-byte array containing the hash value.
' @remarks The TransformFinalBlock must be called before accessing
' this property or an InvalidOperationException will be thrown.
'
Public Property Get Hash() As Byte()
    Hash = mBase.Hash
End Property

''
' Returns the size of the computed hash.
'
' @return Always returns 160 bits.
'
Public Property Get HashSize() As Long
    HashSize = 160
End Property

''
' Returns the size of the input block size allowed.
'
' @returns Always returns 1.
' @remarks When calling the TransformBlock method, the InputCount must be zero or a multiple of this value.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = 1
End Property

''
' Returns the output block size this transform can produce.
'
' @return Always returns 1.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = 1
End Property

''
' Releases any resources.
'
Public Sub Clear()
    Call mBase.Clear
End Sub

''
' Computes a hash value for a source in a single pass.
'
' @param Source This can be either a Byte array or any Stream compatible object.
' @param Index The starting index in the Byte array. This is ignored for Stream objects.
' @param Count The number of bytes in the array to be computed. This is ignored for Stream objects.
' @return A 20-byte hash value.
' @remarks If a Stream object is passed in, then the stream is continuously read and the hash calculated until
' there is no more data left to read from the stream.
'
Public Function ComputeHash(ByRef Source As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    ComputeHash = mBase.ComputeHash(Source, Index, Count)
End Function

''
' Resets the hashing algorithm to start again.
'
Public Sub Initialize()
    Erase mW
    
    mState(0) = &H67452301
    mState(1) = &HEFCDAB89
    mState(2) = &H98BADCFE
    mState(3) = &H10325476
    mState(4) = &HC3D2E1F0
    
    Call mBase.Initialize
End Sub

''
' Computes the hash for the specified block of data.
'
' @param InputBuffer The data to compute the hash from.
' @param InputOffset The starting index in the input data to compute the hash.
' @param InputCount The number of bytes to compute the hash from.
' @param OutputBuffer The array to place the input buffer bytes in.
' @param OutputOffset The starting index to beging copying the bytes to.
' @return The number of bytes processed.
' @remarks The OutputBuffer will contain the same data as the InputBuffer. No hash values are placed in the OutputBuffer.
' <p>If the OutputBuffer is a Null Array, or is the same array as the InputBuffer and the OutputOffset equals the InputOffset, then nothing is copied, however, the hash is still computed
' on the InputBuffer data.</p>
' <p>The TransformFinalBlock needs to be called to finish computing the hash, this method alone cannot compute the entire hash.</p>
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    TransformBlock = mBase.TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

''
' Computes the final hash value.
'
' @param InputBuffer The final block of data to compute the hash for.
' @param InputOffset The index in the InputBuffer to begin computing data.
' @param InputCount The number of bytes to be processed.
' @return Returns the section of the InputBuffer that was processed.
' @remarks This does not return the computed hash value. A copy of the data that was
' process is returned. To retrieve the hash value, call the Hash property.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    TransformFinalBlock = mBase.TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to compare against this instance.
' @return Returns True if the values are the same.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' @return Returns a number identifing this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return Returns a string representation of this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashFinal(ByRef Buffer() As Byte, ByVal BufferCount As Long, ByVal MessageLength As Currency) As Byte()
    Buffer(BufferCount) = &H80
    BufferCount = BufferCount + 1
    
    Dim Size As Long
    Size = IIf(BufferCount > 56, 128, 64)
    
    Dim FinalBuffer() As Byte
    ReDim FinalBuffer(0 To Size - 1)
    Call CopyMemory(FinalBuffer(0), Buffer(0), BufferCount)
    
    Dim i As Long
    With AsOByteArr(MessageLength * 0.0008@)
        For i = 0 To 7
            FinalBuffer(Size - i - 1) = .Bytes(i)
        Next i
    End With
    
    Call Transform(FinalBuffer, 0)
    If Size = 128 Then Call Transform(FinalBuffer, 64)
    
    Dim Ret() As Byte
    ReDim Ret(19)
    For i = 0 To 4
        AsLong(Ret(i * 4)) = SwapEndian(mState(i))
    Next i
    
    HashFinal = Ret
End Function

Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    If InIDE Then
        Call TransformIDE(Bytes, Index)
    Else
        Call TransformNative(Bytes, Index)
    End If
End Sub

''
' Runs in the IDE.
'
Private Sub TransformIDE(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim i As Long
    For i = 0 To 15
        mW(i) = SwapEndian(AsLong(Bytes(Index)))
        Index = Index + 4
    Next i

    For i = 16 To 79
        mW(i) = LRotate(mW(i - 3) Xor mW(i - 8) Xor mW(i - 14) Xor mW(i - 16), 1)
    Next i

    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim k As Long
    Dim t As Long

    a = mState(0)
    b = mState(1)
    c = mState(2)
    d = mState(3)
    e = mState(4)
    
    For i = 0 To 79
        Select Case i
            Case Is < 20
                f = d Xor (b And (c Xor d))
                k = &H5A827999
            Case Is < 40
                f = b Xor c Xor d
                k = &H6ED9EBA1
            Case Is < 60
                f = (b And c) Or (d And (b Or c))
                k = &H8F1BBCDC
            Case Else
                f = b Xor c Xor d
                k = &HCA62C1D6
        End Select
        
        t = Helper.UAdd(Helper.UAdd(Helper.UAdd(Helper.UAdd(LRotate(a, 5), f), e), mW(i)), k)
        e = d
        d = c
        c = LRotate(b, 30)
        b = a
        a = t
    Next i
    
    mState(0) = Helper.UAdd(mState(0), a)
    mState(1) = Helper.UAdd(mState(1), b)
    mState(2) = Helper.UAdd(mState(2), c)
    mState(3) = Helper.UAdd(mState(3), d)
    mState(4) = Helper.UAdd(mState(4), e)
End Sub

''
' Runs when compiled. Must disable Integer overflow checks.
'
Private Sub TransformNative(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim t As Long
    Dim i As Long
    
    For i = 0 To 15
        mW(i) = (Bytes(Index) * &H1000000) Or (Bytes(Index + 1) * &H10000) Or (Bytes(Index + 2) * &H100&) Or Bytes(Index + 3)
        Index = Index + 4
    Next i

    For i = 16 To 79
        t = mW(i - 3) Xor mW(i - 8) Xor mW(i - 14) Xor mW(i - 16)
        mW(i) = (t * &H2) Or ((t And &H80000000) \ &H80000000)
    Next i

    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long

    a = mState(0)
    b = mState(1)
    c = mState(2)
    d = mState(3)
    e = mState(4)
    
    For i = 0 To 19
        t = ((a * &H20&) Or (((a And &HF8000000) \ &H8000000) And &H1F&)) + (d Xor (b And (c Xor d))) + e + mW(i) + &H5A827999
        e = d
        d = c
        c = (b * &H40000000) Or (((b And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)
        b = a
        a = t
    Next i

    For i = 20 To 39
        t = ((a * &H20&) Or (((a And &HF8000000) \ &H8000000) And &H1F&)) + (b Xor c Xor d) + e + mW(i) + &H6ED9EBA1
        e = d
        d = c
        c = (b * &H40000000) Or (((b And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)
        b = a
        a = t
    Next i

    For i = 40 To 59
        t = ((a * &H20&) Or (((a And &HF8000000) \ &H8000000) And &H1F&)) + ((b And c) Or (d And (b Or c))) + e + mW(i) + &H8F1BBCDC
        e = d
        d = c
        c = (b * &H40000000) Or (((b And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)
        b = a
        a = t
    Next i

    For i = 60 To 79
        t = ((a * &H20&) Or (((a And &HF8000000) \ &H8000000) And &H1F&)) + (b Xor c Xor d) + e + mW(i) + &HCA62C1D6
        e = d
        d = c
        c = (b * &H40000000) Or (((b And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF)
        b = a
        a = t
    Next i
    
    mState(0) = mState(0) + a
    mState(1) = mState(1) + b
    mState(2) = mState(2) + c
    mState(3) = mState(3) + d
    mState(4) = mState(4) + e
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = Cor.NewHashAlgorithmBase(Me, True, True, 64)
    Call Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   HashAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HashAlgorithm_CanReuseTransform() As Boolean
    HashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get HashAlgorithm_CanTransformMultipleBlocks() As Boolean
    HashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub HashAlgorithm_Clear()
    Call Clear
End Sub

Private Function HashAlgorithm_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    HashAlgorithm_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function HashAlgorithm_Equals(Value As Variant) As Boolean
    HashAlgorithm_Equals = Equals(Value)
End Function

Private Function HashAlgorithm_GetHashCode() As Long
    HashAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get HashAlgorithm_Hash() As Byte()
    HashAlgorithm_Hash = Hash
End Property

Private Property Get HashAlgorithm_HashSize() As Long
    HashAlgorithm_HashSize = HashSize
End Property

Private Sub HashAlgorithm_Initialize()
    Call Initialize
End Sub

Private Property Get HashAlgorithm_InputBlockSize() As Long
    HashAlgorithm_InputBlockSize = InputBlockSize
End Property

Private Property Get HashAlgorithm_OutputBlockSize() As Long
    HashAlgorithm_OutputBlockSize = OutputBlockSize
End Property

Private Function HashAlgorithm_ToString() As String
    HashAlgorithm_ToString = ToString
End Function

Private Function HashAlgorithm_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    HashAlgorithm_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function HashAlgorithm_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    HashAlgorithm_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ICryptoTransform_CanReuseTransform() As Boolean
    ICryptoTransform_CanReuseTransform = CanReuseTransform
End Property

Private Property Get ICryptoTransform_CanTransformMultipleBlocks() As Boolean
    ICryptoTransform_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Property Get ICryptoTransform_InputBlockSize() As Long
    ICryptoTransform_InputBlockSize = InputBlockSize
End Property

Private Property Get ICryptoTransform_OutputBlockSize() As Long
    ICryptoTransform_OutputBlockSize = OutputBlockSize
End Property

Private Function ICryptoTransform_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ICryptoTransform_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function ICryptoTransform_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    ICryptoTransform_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IHashAlgorithmCallback Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IHashAlgorithm_HashCore(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Call Transform(Bytes, Index)
End Sub

Private Function IHashAlgorithm_HashFinal(ByRef Buffer() As Byte, ByVal ByteCount As Long, ByVal MessageLength As Currency) As Byte()
    IHashAlgorithm_HashFinal = HashFinal(Buffer, ByteCount, MessageLength)
End Function

Private Sub IHashAlgorithm_Initialize()
    Call Me.Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SHA1 Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SHA1_CanReuseTransform() As Boolean
    SHA1_CanReuseTransform = CanReuseTransform
End Property

Private Property Get SHA1_CanTransformMultipleBlocks() As Boolean
    SHA1_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub SHA1_Clear()
    Call Clear
End Sub

Private Function SHA1_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    SHA1_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function SHA1_Equals(Value As Variant) As Boolean
    SHA1_Equals = Equals(Value)
End Function

Private Function SHA1_GetHashCode() As Long
    SHA1_GetHashCode = GetHashCode
End Function

Private Property Get SHA1_Hash() As Byte()
    SHA1_Hash = Hash
End Property

Private Property Get SHA1_HashSize() As Long
    SHA1_HashSize = HashSize
End Property

Private Sub SHA1_Initialize()
    Call Initialize
End Sub

Private Property Get SHA1_InputBlockSize() As Long
    SHA1_InputBlockSize = InputBlockSize
End Property

Private Property Get SHA1_OutputBlockSize() As Long
    SHA1_OutputBlockSize = OutputBlockSize
End Property

Private Function SHA1_ToString() As String
    SHA1_ToString = ToString
End Function

Private Function SHA1_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    SHA1_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function SHA1_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    SHA1_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function



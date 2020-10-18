VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA256Managed"
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
' Module: SHA256Managed
'

''
' Provides services to compute the SHA-256 hash value.
'
' @remarks The SHA-256 algorithm is implemented in Visual Basic.
' <p>Details can be found at http://csrc.nist.gov/CryptoToolkit/tkhash.html</p>
' @see HashAlgorithm
'
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements ICryptoTransform
Implements SHA256
Implements IHashAlgorithm

Private mBase               As HashAlgorithmBase
Private mW(0 To 79)         As Long
Private mK(0 To 63)         As Long
Private mState(0 To 7)      As Long



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
' @return A 32-byte array containing the hash value.
' @remarks The TransformFinalBlock must be called before accessing
' this property or an InvalidOperationException will be thrown.
'
Public Property Get Hash() As Byte()
    Hash = mBase.Hash
End Property

''
' Returns the size of the computed hash.
'
' @return Always returns 256 bits.
'
Public Property Get HashSize() As Long
    HashSize = 256
End Property

''
' Returns the input block size this transform can work on.
'
' @return Always returns 1.
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
    mBase.Clear
End Sub

''
' Computes a hash value for a source in a single pass.
'
' @param Source This can be either a Byte array or any Stream compatible object.
' @param Index The starting index in the Byte array. This is ignored for Stream objects.
' @param Count The number of bytes in the array to be computed. This is ignored for Stream objects.
' @return A 32-byte hash value.
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
    
    mState(0) = &H6A09E667
    mState(1) = &HBB67AE85
    mState(2) = &H3C6EF372
    mState(3) = &HA54FF53A
    mState(4) = &H510E527F
    mState(5) = &H9B05688C
    mState(6) = &H1F83D9AB
    mState(7) = &H5BE0CD19
    
    mBase.Initialize
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
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashFinal(ByRef Buffer() As Byte, ByVal BufferCount As Long, ByVal MessageLength As Currency) As Byte()
    Buffer(BufferCount) = &H80
    BufferCount = BufferCount + 1
    
    Dim Size As Long
    Size = IIf(BufferCount > 56, 128, 64)
    
    Dim FinalBuffer() As Byte
    ReDim FinalBuffer(0 To Size - 1)
    CopyMemory FinalBuffer(0), Buffer(0), BufferCount
    
    Dim i As Long
    With AsOByteArr(MessageLength * 0.0008@)
        For i = 0 To 7
            FinalBuffer(Size - i - 1) = .Bytes(i)
        Next i
    End With
    
    Transform FinalBuffer, 0
    If Size = 128 Then
        Transform FinalBuffer, 64
    End If
    
    Dim Result() As Byte
    ReDim Result(0 To 31)
    For i = 0 To 7
        AsLong(Result(i * 4)) = SwapEndian(mState(i))
    Next i
    
    HashFinal = Result
End Function

#If Release Then
' Must disable Integer overflow checks.
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim i As Long
    For i = 0 To 15
        mW(i) = (Bytes(Index) * &H1000000) Or (Bytes(Index + 1) * &H10000) Or (Bytes(Index + 2) * &H100&) Or Bytes(Index + 3)
        Index = Index + 4
    Next i
    
    Dim t1 As Long
    Dim t2 As Long
    For i = 16 To 63
        t1 = mW(i - 15)
        t2 = mW(i - 2)
        mW(i) = (mW(i - 16) + mW(i - 7)) + _
                (((((t1 And &HFFFFFF80) \ &H80&) And &H1FFFFFF) Or (t1 * &H2000000)) Xor _
                ((((t1 And &HFFFC0000) \ &H40000) And &H3FFF&) Or (t1 * &H4000&)) Xor _
                ((((t1 And &HFFFFFFF8) \ &H8&) And &H1FFFFFFF))) + _
                (((((t2 And &HFFFE0000) \ &H20000) And &H7FFF&) Or (t2 * &H8000&)) Xor _
                ((((t2 And &HFFF80000) \ &H80000) And &H1FFF&) Or (t2 * &H2000&)) Xor _
                ((((t2 And &HFFFFFC00) \ &H400&) And &H3FFFFF)))
    Next i
    
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    
    a = mState(0)
    b = mState(1)
    c = mState(2)
    d = mState(3)
    e = mState(4)
    f = mState(5)
    g = mState(6)
    h = mState(7)
    
    For i = 0 To 63
        t1 = ((a And b) Xor (a And c) Xor (b And c)) + _
             (((((a And &HFFFFFFFC) \ &H4&) And &H3FFFFFFF) Or (a * &H40000000)) Xor _
             ((((a And &HFFFFE000) \ &H2000&) And &H7FFFF) Or (a * &H80000)) Xor _
             ((((a And &HFFC00000) \ &H400000) And &H3FF) Or (a * &H400&)))
        
        
        t2 = (h + mK(i) + mW(i)) + _
             (((((e And &HFFFFFFC0) \ &H40&) And &H3FFFFFF) Or (e * &H4000000)) Xor _
             ((((e And &HFFFFF800) \ &H800&) And &H1FFFFF) Or (e * &H200000)) Xor _
             ((((e And &HFE000000) \ &H2000000) And &H7F) Or (e * &H80&))) + _
             ((e And f) Xor ((Not e) And g))
                
        h = g
        g = f
        f = e
        e = d + t2
        d = c
        c = b
        b = a
        a = t1 + t2
    Next i
    
    mState(0) = mState(0) + a
    mState(1) = mState(1) + b
    mState(2) = mState(2) + c
    mState(3) = mState(3) + d
    mState(4) = mState(4) + e
    mState(5) = mState(5) + f
    mState(6) = mState(6) + g
    mState(7) = mState(7) + h
End Sub
#Else
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    Dim i As Long
    For i = 0 To 15
        mW(i) = SwapEndian(AsLong(Bytes(Index)))
        Index = Index + 4
    Next i
    
    Dim s0 As Long
    Dim s1 As Long
    For i = 16 To 63
        s0 = RRotate(mW(i - 15), 7) Xor RRotate(mW(i - 15), 18) Xor Helper.ShiftRight(mW(i - 15), 3)
        s1 = RRotate(mW(i - 2), 17) Xor RRotate(mW(i - 2), 19) Xor Helper.ShiftRight(mW(i - 2), 10)
        mW(i) = Helper.UAdd(Helper.UAdd(Helper.UAdd(mW(i - 16), s0), s1), mW(i - 7))
    Next i
    
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim t1 As Long
    Dim Maj As Long
    Dim t2 As Long
    Dim Ch As Long
    
    a = mState(0)
    b = mState(1)
    c = mState(2)
    d = mState(3)
    e = mState(4)
    f = mState(5)
    g = mState(6)
    h = mState(7)
    
    For i = 0 To 63
        s0 = RRotate(a, 2) Xor RRotate(a, 13) Xor RRotate(a, 22)
        Maj = ((a And b) Xor (a And c) Xor (b And c))
        t2 = Helper.UAdd(s0, Maj)
        s1 = RRotate(e, 6) Xor RRotate(e, 11) Xor RRotate(e, 25)
        Ch = (e And f) Xor ((Not e) And g)
        t1 = Helper.UAdd(Helper.UAdd(Helper.UAdd(Helper.UAdd(h, s1), Ch), mK(i)), mW(i))
        
        h = g
        g = f
        f = e
        e = Helper.UAdd(d, t1)
        d = c
        c = b
        b = a
        a = Helper.UAdd(t1, t2)
    Next i
    
    mState(0) = Helper.UAdd(mState(0), a)
    mState(1) = Helper.UAdd(mState(1), b)
    mState(2) = Helper.UAdd(mState(2), c)
    mState(3) = Helper.UAdd(mState(3), d)
    mState(4) = Helper.UAdd(mState(4), e)
    mState(5) = Helper.UAdd(mState(5), f)
    mState(6) = Helper.UAdd(mState(6), g)
    mState(7) = Helper.UAdd(mState(7), h)

End Sub
#End If




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    CopyMemory mK(0), ByVal Statics.SHA256.KPtr, 256
    Set mBase = Cor.NewHashAlgorithmBase(Me, True, True, 64)
    Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   HashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HashAlgorithm_CanReuseTransform() As Boolean
    HashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get HashAlgorithm_CanTransformMultipleBlocks() As Boolean
    HashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub HashAlgorithm_Clear()
    Clear
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
    Initialize
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
'   ICryptoTransform
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
'   IHashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub IHashAlgorithm_HashCore(Bytes() As Byte, ByVal Index As Long, ByVal Count As Long)
    Transform Bytes, Index
End Sub

Private Function IHashAlgorithm_HashFinal(ByRef Buffer() As Byte, ByVal ByteCount As Long, ByVal MessageLength As Currency) As Byte()
    IHashAlgorithm_HashFinal = HashFinal(Buffer, ByteCount, MessageLength)
End Function

Private Sub IHashAlgorithm_Initialize()
    Initialize
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SHA256
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SHA256_CanReuseTransform() As Boolean
    SHA256_CanReuseTransform = CanReuseTransform
End Property

Private Property Get SHA256_CanTransformMultipleBlocks() As Boolean
    SHA256_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub SHA256_Clear()
    Clear
End Sub

Private Function SHA256_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    SHA256_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function SHA256_Equals(Value As Variant) As Boolean
    SHA256_Equals = Equals(Value)
End Function

Private Function SHA256_GetHashCode() As Long
    SHA256_GetHashCode = GetHashCode
End Function

Private Property Get SHA256_Hash() As Byte()
    SHA256_Hash = Hash
End Property

Private Property Get SHA256_HashSize() As Long
    SHA256_HashSize = HashSize
End Property

Private Sub SHA256_Initialize()
    Initialize
End Sub

Private Property Get SHA256_InputBlockSize() As Long
    SHA256_InputBlockSize = InputBlockSize
End Property

Private Property Get SHA256_OutputBlockSize() As Long
    SHA256_OutputBlockSize = OutputBlockSize
End Property

Private Function SHA256_ToString() As String
    SHA256_ToString = ToString
End Function

Private Function SHA256_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    SHA256_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function SHA256_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    SHA256_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RIPEMD160Managed"
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
' Module: RIPEMD160Managed
'

''
' Represents an implementation of the RIPEMD160 hashing algorithm.
'
' @remarks Details about the hash algorithm can be found here: http://homes.esat.kuleuven.be/~bosselae/ripemd160.html
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements RIPEMD160
Implements ICryptoTransform
Implements IHashAlgorithm

Private mBase               As HashAlgorithmBase
Private mState(0 To 4)      As Long
Private mX(0 To 15)         As Long


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
    mBase.Clear
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
    mState(0) = &H67452301
    mState(1) = &HEFCDAB89
    mState(2) = &H98BADCFE
    mState(3) = &H10325476
    mState(4) = &HC3D2E1F0
    
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
    CopyMemory FinalBuffer(Size - 8), MessageLength * 0.0008@, 8
    Transform FinalBuffer, 0
    If Size = 128 Then
        Transform FinalBuffer, 64
    End If
    
    Dim Result(0 To 19) As Byte
    CopyMemory Result(0), mState(0), 20

    HashFinal = Result
End Function

#If Release Then
' Must disable Integer overflow checks.
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    CopyMemory mX(0), Bytes(Index), 64
    
    Dim aa As Long
    Dim bb As Long
    Dim dd As Long
    Dim cc As Long
    Dim ee As Long
    Dim aaa As Long
    Dim bbb As Long
    Dim ccc As Long
    Dim ddd As Long
    Dim eee As Long
    
    aa = mState(0)
    bb = mState(1)
    cc = mState(2)
    dd = mState(3)
    ee = mState(4)
    aaa = aa
    bbb = bb
    ccc = cc
    ddd = dd
    eee = ee
    
    'fnFF(aa, bb, cc, dd, ee, mX(0), 11)
    aa = aa + (bb Xor cc Xor dd) + mX(0)
    aa = ((aa * &H800&) Or (((aa And &HFFE00000) \ &H200000) And &H7FF&)) + ee
    cc = (cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&)
    
    'fnFF ee, aa, bb, cc, dd, mX(1), 14)
    ee = ee + (aa Xor bb Xor cc) + mX(1)
    ee = ((ee * &H4000&) Or (((ee And &HFFFC0000) \ &H40000) And &H3FFF&)) + dd
    bb = (bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&)
    
    'fnFF(dd, ee, aa, bb, cc, mX(2), 15)
    dd = dd + (ee Xor aa Xor bb) + mX(2)
    dd = ((dd * &H8000&) Or (((dd And &HFFFE0000) \ &H20000) And &H7FFF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(cc, dd, ee, aa, bb, mX(3), 12)
    cc = cc + (dd Xor ee Xor aa) + mX(3)
    cc = ((cc * &H1000&) Or (((cc And &HFFF00000) \ &H100000) And &HFFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
        
    'fnFF(bb, cc, dd, ee, aa, mX(4), 5)
    bb = bb + (cc Xor dd Xor ee) + mX(4)
    bb = ((bb * &H20&) Or (((bb And &HF8000000) \ &H8000000) And &H1F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(aa, bb, cc, dd, ee, mX(5), 8)
    aa = aa + (bb Xor cc Xor dd) + mX(5)
    aa = ((aa * &H100&) Or (((aa And &HFF000000) \ &H1000000) And &HFF&)) + ee
    cc = (cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&)
    
    'fnFF(ee, aa, bb, cc, dd, mX(6), 7)
    ee = ee + (aa Xor bb Xor cc) + mX(6)
    ee = ((ee * &H80&) Or (((ee And &HFE000000) \ &H2000000) And &H7F&)) + dd
    bb = (bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&)
        
    'fnFF(dd, ee, aa, bb, cc, mX(7), 9)
    dd = dd + (ee Xor aa Xor bb) + mX(7)
    dd = ((dd * &H200&) Or (((dd And &HFF800000) \ &H800000) And &H1FF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(cc, dd, ee, aa, bb, mX(8), 11)
    cc = cc + (dd Xor ee Xor aa) + mX(8)
    cc = ((cc * &H800&) Or (((cc And &HFFE00000) \ &H200000) And &H7FF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFF(bb, cc, dd, ee, aa, mX(9), 13)
    bb = bb + (cc Xor dd Xor ee) + mX(9)
    bb = ((bb * &H2000&) Or (((bb And &HFFF80000) \ &H80000) And &H1FFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(aa, bb, cc, dd, ee, mX(10), 14)
    aa = aa + (bb Xor cc Xor dd) + mX(10)
    aa = ((aa * &H4000&) Or (((aa And &HFFFC0000) \ &H40000) And &H3FFF&)) + ee
    cc = (cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&)
    
    'fnFF(ee, aa, bb, cc, dd, mX(11), 15)
    ee = ee + (aa Xor bb Xor cc) + mX(11)
    ee = ((ee * &H8000&) Or (((ee And &HFFFE0000) \ &H20000) And &H7FFF&)) + dd
    bb = (bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&)
    
    'fnFF(dd, ee, aa, bb, cc, mX(12), 6)
    dd = dd + (ee Xor aa Xor bb) + mX(12)
    dd = ((dd * &H40&) Or (((dd And &HFC000000) \ &H4000000) And &H3F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(cc, dd, ee, aa, bb, mX(13), 7)
    cc = cc + (dd Xor ee Xor aa) + mX(13)
    cc = ((cc * &H80&) Or (((cc And &HFE000000) \ &H2000000) And &H7F&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(bb, cc, dd, ee, aa, mX(14), 9)
    bb = bb + (cc Xor dd Xor ee) + mX(14)
    bb = ((bb * &H200&) Or (((bb And &HFF800000) \ &H800000) And &H1FF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFF(aa, bb, cc, dd, ee, mX(15), 8)
    aa = aa + (bb Xor cc Xor dd) + mX(15)
    aa = ((aa * &H100&) Or (((aa And &HFF000000) \ &H1000000) And &HFF&)) + ee
    cc = (cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&)

    'fnGG(ee, aa, bb, cc, dd, mX(7), 7)
    ee = ee + ((aa And bb) Or ((Not aa) And cc)) + mX(7) + &H5A827999
    ee = ((ee * &H80&) Or (((ee And &HFE000000) \ &H2000000) And &H7F&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(dd, ee, aa, bb, cc, mX(4), 6)
    dd = dd + ((ee And aa) Or ((Not ee) And bb)) + mX(4) + &H5A827999
    dd = ((dd * &H40&) Or (((dd And &HFC000000) \ &H4000000) And &H3F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(cc, dd, ee, aa, bb, mX(13), 8)
    cc = cc + ((dd And ee) Or ((Not dd) And aa)) + mX(13) + &H5A827999
    cc = ((cc * &H100&) Or (((cc And &HFF000000) \ &H1000000) And &HFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(bb, cc, dd, ee, aa, mX(1), 13)
    bb = bb + ((cc And dd) Or ((Not cc) And ee)) + mX(1) + &H5A827999
    bb = ((bb * &H2000&) Or (((bb And &HFFF80000) \ &H80000) And &H1FFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(aa, bb, cc, dd, ee, mX(10), 11)
    aa = aa + ((bb And cc) Or ((Not bb) And dd)) + mX(10) + &H5A827999
    aa = ((aa * &H800&) Or (((aa And &HFFE00000) \ &H200000) And &H7FF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(ee, aa, bb, cc, dd, mX(6), 9)
    ee = ee + ((aa And bb) Or ((Not aa) And cc)) + mX(6) + &H5A827999
    ee = ((ee * &H200&) Or (((ee And &HFF800000) \ &H800000) And &H1FF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(dd, ee, aa, bb, cc, mX(15), 7)
    dd = dd + ((ee And aa) Or ((Not ee) And bb)) + mX(15) + &H5A827999
    dd = ((dd * &H80&) Or (((dd And &HFE000000) \ &H2000000) And &H7F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(cc, dd, ee, aa, bb, mX(3), 15)
    cc = cc + ((dd And ee) Or ((Not dd) And aa)) + mX(3) + &H5A827999
    cc = ((cc * &H8000&) Or (((cc And &HFFFE0000) \ &H20000) And &H7FFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(bb, cc, dd, ee, aa, mX(12), 7)
    bb = bb + ((cc And dd) Or ((Not cc) And ee)) + mX(12) + &H5A827999
    bb = ((bb * &H80&) Or (((bb And &HFE000000) \ &H2000000) And &H7F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGG(aa, bb, cc, dd, ee, mX(0), 12)
    aa = aa + ((bb And cc) Or ((Not bb) And dd)) + mX(0) + &H5A827999
    aa = ((aa * &H1000&) Or (((aa And &HFFF00000) \ &H100000) And &HFFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(ee, aa, bb, cc, dd, mX(9), 15)
    ee = ee + ((aa And bb) Or ((Not aa) And cc)) + mX(9) + &H5A827999
    ee = ((ee * &H8000&) Or (((ee And &HFFFE0000) \ &H20000) And &H7FFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(dd, ee, aa, bb, cc, mX(5), 9)
    dd = dd + ((ee And aa) Or ((Not ee) And bb)) + mX(5) + &H5A827999
    dd = ((dd * &H200&) Or (((dd And &HFF800000) \ &H800000) And &H1FF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(cc, dd, ee, aa, bb, mX(2), 11)
    cc = cc + ((dd And ee) Or ((Not dd) And aa)) + mX(2) + &H5A827999
    cc = ((cc * &H800&) Or (((cc And &HFFE00000) \ &H200000) And &H7FF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(bb, cc, dd, ee, aa, mX(14), 7)
    bb = bb + ((cc And dd) Or ((Not cc) And ee)) + mX(14) + &H5A827999
    bb = ((bb * &H80&) Or (((bb And &HFE000000) \ &H2000000) And &H7F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(aa, bb, cc, dd, ee, mX(11), 13)
    aa = aa + ((bb And cc) Or ((Not bb) And dd)) + mX(11) + &H5A827999
    aa = ((aa * &H2000&) Or (((aa And &HFFF80000) \ &H80000) And &H1FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGG(ee, aa, bb, cc, dd, mX(8), 12)
    ee = ee + ((aa And bb) Or ((Not aa) And cc)) + mX(8) + &H5A827999
    ee = ((ee * &H1000&) Or (((ee And &HFFF00000) \ &H100000) And &HFFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(dd, ee, aa, bb, cc, mX(3), 11)
    dd = dd + ((ee Or (Not aa)) Xor bb) + mX(3) + &H6ED9EBA1
    dd = ((dd * &H800&) Or (((dd And &HFFE00000) \ &H200000) And &H7FF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(cc, dd, ee, aa, bb, mX(10), 13)
    cc = cc + ((dd Or (Not ee)) Xor aa) + mX(10) + &H6ED9EBA1
    cc = ((cc * &H2000&) Or (((cc And &HFFF80000) \ &H80000) And &H1FFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(bb, cc, dd, ee, aa, mX(14), 6)
    bb = bb + ((cc Or (Not dd)) Xor ee) + mX(14) + &H6ED9EBA1
    bb = ((bb * &H40&) Or (((bb And &HFC000000) \ &H4000000) And &H3F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(aa, bb, cc, dd, ee, mX(4), 7)
    aa = aa + ((bb Or (Not cc)) Xor dd) + mX(4) + &H6ED9EBA1
    aa = ((aa * &H80&) Or (((aa And &HFE000000) \ &H2000000) And &H7F&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(ee, aa, bb, cc, dd, mX(9), 14)
    ee = ee + ((aa Or (Not bb)) Xor cc) + mX(9) + &H6ED9EBA1
    ee = ((ee * &H4000&) Or (((ee And &HFFFC0000) \ &H40000) And &H3FFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(dd, ee, aa, bb, cc, mX(15), 9)
    dd = dd + ((ee Or (Not aa)) Xor bb) + mX(15) + &H6ED9EBA1
    dd = ((dd * &H200&) Or (((dd And &HFF800000) \ &H800000) And &H1FF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(cc, dd, ee, aa, bb, mX(8), 13)
    cc = cc + ((dd Or (Not ee)) Xor aa) + mX(8) + &H6ED9EBA1
    cc = ((cc * &H2000&) Or (((cc And &HFFF80000) \ &H80000) And &H1FFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(bb, cc, dd, ee, aa, mX(1), 15)
    bb = bb + ((cc Or (Not dd)) Xor ee) + mX(1) + &H6ED9EBA1
    bb = ((bb * &H8000&) Or (((bb And &HFFFE0000) \ &H20000) And &H7FFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(aa, bb, cc, dd, ee, mX(2), 14)
    aa = aa + ((bb Or (Not cc)) Xor dd) + mX(2) + &H6ED9EBA1
    aa = ((aa * &H4000&) Or (((aa And &HFFFC0000) \ &H40000) And &H3FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHH(ee, aa, bb, cc, dd, mX(7), 8)
    ee = ee + ((aa Or (Not bb)) Xor cc) + mX(7) + &H6ED9EBA1
    ee = ((ee * &H100&) Or (((ee And &HFF000000) \ &H1000000) And &HFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(dd, ee, aa, bb, cc, mX(0), 13)
    dd = dd + ((ee Or (Not aa)) Xor bb) + mX(0) + &H6ED9EBA1
    dd = ((dd * &H2000&) Or (((dd And &HFFF80000) \ &H80000) And &H1FFF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(cc, dd, ee, aa, bb, mX(6), 6)
    cc = cc + ((dd Or (Not ee)) Xor aa) + mX(6) + &H6ED9EBA1
    cc = ((cc * &H40&) Or (((cc And &HFC000000) \ &H4000000) And &H3F&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(bb, cc, dd, ee, aa, mX(13), 5)
    bb = bb + ((cc Or (Not dd)) Xor ee) + mX(13) + &H6ED9EBA1
    bb = ((bb * &H20&) Or (((bb And &HF8000000) \ &H8000000) And &H1F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(aa, bb, cc, dd, ee, mX(11), 12)
    aa = aa + ((bb Or (Not cc)) Xor dd) + mX(11) + &H6ED9EBA1
    aa = ((aa * &H1000&) Or (((aa And &HFFF00000) \ &H100000) And &HFFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(ee, aa, bb, cc, dd, mX(5), 7)
    ee = ee + ((aa Or (Not bb)) Xor cc) + mX(5) + &H6ED9EBA1
    ee = ((ee * &H80&) Or (((ee And &HFE000000) \ &H2000000) And &H7F&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHH(dd, ee, aa, bb, cc, mX(12), 5)
    dd = dd + ((ee Or (Not aa)) Xor bb) + mX(12) + &H6ED9EBA1
    dd = ((dd * &H20&) Or (((dd And &HF8000000) \ &H8000000) And &H1F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
        
    'fnII(cc, dd, ee, aa, bb, mX(1), 11)
    cc = cc + ((dd And aa) Or (ee And (Not aa))) + mX(1) + &H8F1BBCDC
    cc = ((cc * &H800&) Or (((cc And &HFFE00000) \ &H200000) And &H7FF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(bb, cc, dd, ee, aa, mX(9), 12)
    bb = bb + ((cc And ee) Or (dd And (Not ee))) + mX(9) + &H8F1BBCDC
    bb = ((bb * &H1000&) Or (((bb And &HFFF00000) \ &H100000) And &HFFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(aa, bb, cc, dd, ee, mX(11), 14)
    aa = aa + ((bb And dd) Or (cc And (Not dd))) + mX(11) + &H8F1BBCDC
    aa = ((aa * &H4000&) Or (((aa And &HFFFC0000) \ &H40000) And &H3FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
        
    'fnII(ee, aa, bb, cc, dd, mX(10), 15)
    ee = ee + ((aa And cc) Or (bb And (Not cc))) + mX(10) + &H8F1BBCDC
    ee = ((ee * &H8000&) Or (((ee And &HFFFE0000) \ &H20000) And &H7FFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(dd, ee, aa, bb, cc, mX(0), 14)
    dd = dd + ((ee And bb) Or (aa And (Not bb))) + mX(0) + &H8F1BBCDC
    dd = ((dd * &H4000&) Or (((dd And &HFFFC0000) \ &H40000) And &H3FFF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnII(cc, dd, ee, aa, bb, mX(8), 15)
    cc = cc + ((dd And aa) Or (ee And (Not aa))) + mX(8) + &H8F1BBCDC
    cc = ((cc * &H8000&) Or (((cc And &HFFFE0000) \ &H20000) And &H7FFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(bb, cc, dd, ee, aa, mX(12), 9)
    bb = bb + ((cc And ee) Or (dd And (Not ee))) + mX(12) + &H8F1BBCDC
    bb = ((bb * &H200&) Or (((bb And &HFF800000) \ &H800000) And &H1FF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(aa, bb, cc, dd, ee, mX(4), 8)
    aa = aa + ((bb And dd) Or (cc And (Not dd))) + mX(4) + &H8F1BBCDC
    aa = ((aa * &H100&) Or (((aa And &HFF000000) \ &H1000000) And &HFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(ee, aa, bb, cc, dd, mX(13), 9)
    ee = ee + ((aa And cc) Or (bb And (Not cc))) + mX(13) + &H8F1BBCDC
    ee = ((ee * &H200&) Or (((ee And &HFF800000) \ &H800000) And &H1FF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(dd, ee, aa, bb, cc, mX(3), 14)
    dd = dd + ((ee And bb) Or (aa And (Not bb))) + mX(3) + &H8F1BBCDC
    dd = ((dd * &H4000&) Or (((dd And &HFFFC0000) \ &H40000) And &H3FFF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(cc, dd, ee, aa, bb, mX(7), 5)
    cc = cc + ((dd And aa) Or (ee And (Not aa))) + mX(7) + &H8F1BBCDC
    cc = ((cc * &H20&) Or (((cc And &HF8000000) \ &H8000000) And &H1F&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(bb, cc, dd, ee, aa, mX(15), 6)
    bb = bb + ((cc And ee) Or (dd And (Not ee))) + mX(15) + &H8F1BBCDC
    bb = ((bb * &H40&) Or (((bb And &HFC000000) \ &H4000000) And &H3F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(aa, bb, cc, dd, ee, mX(14), 8)
    aa = aa + ((bb And dd) Or (cc And (Not dd))) + mX(14) + &H8F1BBCDC
    aa = ((aa * &H100&) Or (((aa And &HFF000000) \ &H1000000) And &HFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
        
    'fnII(ee, aa, bb, cc, dd, mX(5), 6)
    ee = ee + ((aa And cc) Or (bb And (Not cc))) + mX(5) + &H8F1BBCDC
    ee = ((ee * &H40&) Or (((ee And &HFC000000) \ &H4000000) And &H3F&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(dd, ee, aa, bb, cc, mX(6), 5)
    dd = dd + ((ee And bb) Or (aa And (Not bb))) + mX(6) + &H8F1BBCDC
    dd = ((dd * &H20&) Or (((dd And &HF8000000) \ &H8000000) And &H1F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnII(cc, dd, ee, aa, bb, mX(2), 12)
    cc = cc + ((dd And aa) Or (ee And (Not aa))) + mX(2) + &H8F1BBCDC
    cc = ((cc * &H1000&) Or (((cc And &HFFF00000) \ &H100000) And &HFFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(bb, cc, dd, ee, aa, mX(4), 9)
    bb = bb + (cc Xor (dd Or (Not ee))) + mX(4) + &HA953FD4E
    bb = ((bb * &H200&) Or (((bb And &HFF800000) \ &H800000) And &H1FF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(aa, bb, cc, dd, ee, mX(0), 15)
    aa = aa + (bb Xor (cc Or (Not dd))) + mX(0) + &HA953FD4E
    aa = ((aa * &H8000&) Or (((aa And &HFFFE0000) \ &H20000) And &H7FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJ(ee, aa, bb, cc, dd, mX(5), 5)
    ee = ee + (aa Xor (bb Or (Not cc))) + mX(5) + &HA953FD4E
    ee = ((ee * &H20&) Or (((ee And &HF8000000) \ &H8000000) And &H1F&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJ(dd, ee, aa, bb, cc, mX(9), 11)
    dd = dd + (ee Xor (aa Or (Not bb))) + mX(9) + &HA953FD4E
    dd = ((dd * &H800&) Or (((dd And &HFFE00000) \ &H200000) And &H7FF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJ(cc, dd, ee, aa, bb, mX(7), 6)
    cc = cc + (dd Xor (ee Or (Not aa))) + mX(7) + &HA953FD4E
    cc = ((cc * &H40&) Or (((cc And &HFC000000) \ &H4000000) And &H3F&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJ(bb, cc, dd, ee, aa, mX(12), 8)
    bb = bb + (cc Xor (dd Or (Not ee))) + mX(12) + &HA953FD4E
    bb = ((bb * &H100&) Or (((bb And &HFF000000) \ &H1000000) And &HFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJ(aa, bb, cc, dd, ee, mX(2), 13)
    aa = aa + (bb Xor (cc Or (Not dd))) + mX(2) + &HA953FD4E
    aa = ((aa * &H2000&) Or (((aa And &HFFF80000) \ &H80000) And &H1FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(ee, aa, bb, cc, dd, mX(10), 12)
    ee = ee + (aa Xor (bb Or (Not cc))) + mX(10) + &HA953FD4E
    ee = ((ee * &H1000&) Or (((ee And &HFFF00000) \ &H100000) And &HFFF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(dd, ee, aa, bb, cc, mX(14), 5)
    dd = dd + (ee Xor (aa Or (Not bb))) + mX(14) + &HA953FD4E
    dd = ((dd * &H20&) Or (((dd And &HF8000000) \ &H8000000) And &H1F&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(cc, dd, ee, aa, bb, mX(1), 12)
    cc = cc + (dd Xor (ee Or (Not aa))) + mX(1) + &HA953FD4E
    cc = ((cc * &H1000&) Or (((cc And &HFFF00000) \ &H100000) And &HFFF&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(bb, cc, dd, ee, aa, mX(3), 13)
    bb = bb + (cc Xor (dd Or (Not ee))) + mX(3) + &HA953FD4E
    bb = ((bb * &H2000&) Or (((bb And &HFFF80000) \ &H80000) And &H1FFF&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(aa, bb, cc, dd, ee, mX(8), 14)
    aa = aa + (bb Xor (cc Or (Not dd))) + mX(8) + &HA953FD4E
    aa = ((aa * &H4000&) Or (((aa And &HFFFC0000) \ &H40000) And &H3FFF&)) + ee
    cc = ((cc * &H400&) Or (((cc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(ee, aa, bb, cc, dd, mX(11), 11)
    ee = ee + (aa Xor (bb Or (Not cc))) + mX(11) + &HA953FD4E
    ee = ((ee * &H800&) Or (((ee And &HFFE00000) \ &H200000) And &H7FF&)) + dd
    bb = ((bb * &H400&) Or (((bb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(dd, ee, aa, bb, cc, mX(6), 8)
    dd = dd + (ee Xor (aa Or (Not bb))) + mX(6) + &HA953FD4E
    dd = ((dd * &H100&) Or (((dd And &HFF000000) \ &H1000000) And &HFF&)) + cc
    aa = ((aa * &H400&) Or (((aa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(cc, dd, ee, aa, bb, mX(15), 5)
    cc = cc + (dd Xor (ee Or (Not aa))) + mX(15) + &HA953FD4E
    cc = ((cc * &H20&) Or (((cc And &HF8000000) \ &H8000000) And &H1F&)) + bb
    ee = ((ee * &H400&) Or (((ee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJ(bb, cc, dd, ee, aa, mX(13), 6)
    bb = bb + (cc Xor (dd Or (Not ee))) + mX(13) + &HA953FD4E
    bb = ((bb * &H40&) Or (((bb And &HFC000000) \ &H4000000) And &H3F&)) + aa
    dd = ((dd * &H400&) Or (((dd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(aaa, bbb, ccc, ddd, eee, mX(5), 8)
    aaa = aaa + (bbb Xor (ccc Or (Not ddd))) + mX(5) + &H50A28BE6
    aaa = ((aaa * &H100&) Or (((aaa And &HFF000000) \ &H1000000) And &HFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJJ(eee, aaa, bbb, ccc, ddd, mX(14), 9)
    eee = eee + (aaa Xor (bbb Or (Not ccc))) + mX(14) + &H50A28BE6
    eee = ((eee * &H200&) Or (((eee And &HFF800000) \ &H800000) And &H1FF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ddd, eee, aaa, bbb, ccc, mX(7), 9)
    ddd = ddd + (eee Xor (aaa Or (Not bbb))) + mX(7) + &H50A28BE6
    ddd = ((ddd * &H200&) Or (((ddd And &HFF800000) \ &H800000) And &H1FF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ccc, ddd, eee, aaa, bbb, mX(0), 11)
    ccc = ccc + (ddd Xor (eee Or (Not aaa))) + mX(0) + &H50A28BE6
    ccc = ((ccc * &H800&) Or (((ccc And &HFFE00000) \ &H200000) And &H7FF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(bbb, ccc, ddd, eee, aaa, mX(9), 13)
    bbb = bbb + (ccc Xor (ddd Or (Not eee))) + mX(9) + &H50A28BE6
    bbb = ((bbb * &H2000&) Or (((bbb And &HFFF80000) \ &H80000) And &H1FFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(aaa, bbb, ccc, ddd, eee, mX(2), 15)
    aaa = aaa + (bbb Xor (ccc Or (Not ddd))) + mX(2) + &H50A28BE6
    aaa = ((aaa * &H8000&) Or (((aaa And &HFFFE0000) \ &H20000) And &H7FFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(eee, aaa, bbb, ccc, ddd, mX(11), 15)
    eee = eee + (aaa Xor (bbb Or (Not ccc))) + mX(11) + &H50A28BE6
    eee = ((eee * &H8000&) Or (((eee And &HFFFE0000) \ &H20000) And &H7FFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ddd, eee, aaa, bbb, ccc, mX(4), 5)
    ddd = ddd + (eee Xor (aaa Or (Not bbb))) + mX(4) + &H50A28BE6
    ddd = ((ddd * &H20&) Or (((ddd And &HF8000000) \ &H8000000) And &H1F&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ccc, ddd, eee, aaa, bbb, mX(13), 7)
    ccc = ccc + (ddd Xor (eee Or (Not aaa))) + mX(13) + &H50A28BE6
    ccc = ((ccc * &H80&) Or (((ccc And &HFE000000) \ &H2000000) And &H7F&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJJ(bbb, ccc, ddd, eee, aaa, mX(6), 7)
    bbb = bbb + (ccc Xor (ddd Or (Not eee))) + mX(6) + &H50A28BE6
    bbb = ((bbb * &H80&) Or (((bbb And &HFE000000) \ &H2000000) And &H7F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(aaa, bbb, ccc, ddd, eee, mX(15), 8)
    aaa = aaa + (bbb Xor (ccc Or (Not ddd))) + mX(15) + &H50A28BE6
    aaa = ((aaa * &H100&) Or (((aaa And &HFF000000) \ &H1000000) And &HFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(eee, aaa, bbb, ccc, ddd, mX(8), 11)
    eee = eee + (aaa Xor (bbb Or (Not ccc))) + mX(8) + &H50A28BE6
    eee = ((eee * &H800&) Or (((eee And &HFFE00000) \ &H200000) And &H7FF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ddd, eee, aaa, bbb, ccc, mX(1), 14)
    ddd = ddd + (eee Xor (aaa Or (Not bbb))) + mX(1) + &H50A28BE6
    ddd = ((ddd * &H4000&) Or (((ddd And &HFFFC0000) \ &H40000) And &H3FFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(ccc, ddd, eee, aaa, bbb, mX(10), 14)
    ccc = ccc + (ddd Xor (eee Or (Not aaa))) + mX(10) + &H50A28BE6
    ccc = ((ccc * &H4000&) Or (((ccc And &HFFFC0000) \ &H40000) And &H3FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnJJJ(bbb, ccc, ddd, eee, aaa, mX(3), 12)
    bbb = bbb + (ccc Xor (ddd Or (Not eee))) + mX(3) + &H50A28BE6
    bbb = ((bbb * &H1000&) Or (((bbb And &HFFF00000) \ &H100000) And &HFFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnJJJ(aaa, bbb, ccc, ddd, eee, mX(12), 6)
    aaa = aaa + (bbb Xor (ccc Or (Not ddd))) + mX(12) + &H50A28BE6
    aaa = ((aaa * &H40&) Or (((aaa And &HFC000000) \ &H4000000) And &H3F&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(eee, aaa, bbb, ccc, ddd, mX(6), 9)
    eee = eee + ((aaa And ccc) Or (bbb And (Not ccc))) + mX(6) + &H5C4DD124
    eee = ((eee * &H200&) Or (((eee And &HFF800000) \ &H800000) And &H1FF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnIII(ddd, eee, aaa, bbb, ccc, mX(11), 13)
    ddd = ddd + ((eee And bbb) Or (aaa And (Not bbb))) + mX(11) + &H5C4DD124
    ddd = ((ddd * &H2000&) Or (((ddd And &HFFF80000) \ &H80000) And &H1FFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnIII(ccc, ddd, eee, aaa, bbb, mX(3), 15)
    ccc = ccc + ((ddd And aaa) Or (eee And (Not aaa))) + mX(3) + &H5C4DD124
    ccc = ((ccc * &H8000&) Or (((ccc And &HFFFE0000) \ &H20000) And &H7FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(bbb, ccc, ddd, eee, aaa, mX(7), 7)
    bbb = bbb + ((ccc And eee) Or (ddd And (Not eee))) + mX(7) + &H5C4DD124
    bbb = ((bbb * &H80&) Or (((bbb And &HFE000000) \ &H2000000) And &H7F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(aaa, bbb, ccc, ddd, eee, mX(0), 12)
    aaa = aaa + ((bbb And ddd) Or (ccc And (Not ddd))) + mX(0) + &H5C4DD124
    aaa = ((aaa * &H1000&) Or (((aaa And &HFFF00000) \ &H100000) And &HFFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(eee, aaa, bbb, ccc, ddd, mX(13), 8)
    eee = eee + ((aaa And ccc) Or (bbb And (Not ccc))) + mX(13) + &H5C4DD124
    eee = ((eee * &H100&) Or (((eee And &HFF000000) \ &H1000000) And &HFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnIII(ddd, eee, aaa, bbb, ccc, mX(5), 9)
    ddd = ddd + ((eee And bbb) Or (aaa And (Not bbb))) + mX(5) + &H5C4DD124
    ddd = ((ddd * &H200&) Or (((ddd And &HFF800000) \ &H800000) And &H1FF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnIII(ccc, ddd, eee, aaa, bbb, mX(10), 11)
    ccc = ccc + ((ddd And aaa) Or (eee And (Not aaa))) + mX(10) + &H5C4DD124
    ccc = ((ccc * &H800&) Or (((ccc And &HFFE00000) \ &H200000) And &H7FF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(bbb, ccc, ddd, eee, aaa, mX(14), 7)
    bbb = bbb + ((ccc And eee) Or (ddd And (Not eee))) + mX(14) + &H5C4DD124
    bbb = ((bbb * &H80&) Or (((bbb And &HFE000000) \ &H2000000) And &H7F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(aaa, bbb, ccc, ddd, eee, mX(15), 7)
    aaa = aaa + ((bbb And ddd) Or (ccc And (Not ddd))) + mX(15) + &H5C4DD124
    aaa = ((aaa * &H80&) Or (((aaa And &HFE000000) \ &H2000000) And &H7F&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnIII(eee, aaa, bbb, ccc, ddd, mX(8), 12)
    eee = eee + ((aaa And ccc) Or (bbb And (Not ccc))) + mX(8) + &H5C4DD124
    eee = ((eee * &H1000&) Or (((eee And &HFFF00000) \ &H100000) And &HFFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(ddd, eee, aaa, bbb, ccc, mX(12), 7)
    ddd = ddd + ((eee And bbb) Or (aaa And (Not bbb))) + mX(12) + &H5C4DD124
    ddd = ((ddd * &H80&) Or (((ddd And &HFE000000) \ &H2000000) And &H7F&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(ccc, ddd, eee, aaa, bbb, mX(4), 6)
    ccc = ccc + ((ddd And aaa) Or (eee And (Not aaa))) + mX(4) + &H5C4DD124
    ccc = ((ccc * &H40&) Or (((ccc And &HFC000000) \ &H4000000) And &H3F&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(bbb, ccc, ddd, eee, aaa, mX(9), 15)
    bbb = bbb + ((ccc And eee) Or (ddd And (Not eee))) + mX(9) + &H5C4DD124
    bbb = ((bbb * &H8000&) Or (((bbb And &HFFFE0000) \ &H20000) And &H7FFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(aaa, bbb, ccc, ddd, eee, mX(1), 13)
    aaa = aaa + ((bbb And ddd) Or (ccc And (Not ddd))) + mX(1) + &H5C4DD124
    aaa = ((aaa * &H2000&) Or (((aaa And &HFFF80000) \ &H80000) And &H1FFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnIII(eee, aaa, bbb, ccc, ddd, mX(2), 11)
    eee = eee + ((aaa And ccc) Or (bbb And (Not ccc))) + mX(2) + &H5C4DD124
    eee = ((eee * &H800&) Or (((eee And &HFFE00000) \ &H200000) And &H7FF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(ddd, eee, aaa, bbb, ccc, mX(15), 9)
    ddd = ddd + ((eee Or (Not aaa)) Xor bbb) + mX(15) + &H6D703EF3
    ddd = ((ddd * &H200&) Or (((ddd And &HFF800000) \ &H800000) And &H1FF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(ccc, ddd, eee, aaa, bbb, mX(5), 7)
    ccc = ccc + ((ddd Or (Not eee)) Xor aaa) + mX(5) + &H6D703EF3
    ccc = ((ccc * &H80&) Or (((ccc And &HFE000000) \ &H2000000) And &H7F&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(bbb, ccc, ddd, eee, aaa, mX(1), 15)
    bbb = bbb + ((ccc Or (Not ddd)) Xor eee) + mX(1) + &H6D703EF3
    bbb = ((bbb * &H8000&) Or (((bbb And &HFFFE0000) \ &H20000) And &H7FFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(aaa, bbb, ccc, ddd, eee, mX(3), 11)
    aaa = aaa + ((bbb Or (Not ccc)) Xor ddd) + mX(3) + &H6D703EF3
    aaa = ((aaa * &H800&) Or (((aaa And &HFFE00000) \ &H200000) And &H7FF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(eee, aaa, bbb, ccc, ddd, mX(7), 8)
    eee = eee + ((aaa Or (Not bbb)) Xor ccc) + mX(7) + &H6D703EF3
    eee = ((eee * &H100&) Or (((eee And &HFF000000) \ &H1000000) And &HFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(ddd, eee, aaa, bbb, ccc, mX(14), 6)
    ddd = ddd + ((eee Or (Not aaa)) Xor bbb) + mX(14) + &H6D703EF3
    ddd = ((ddd * &H40&) Or (((ddd And &HFC000000) \ &H4000000) And &H3F&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(ccc, ddd, eee, aaa, bbb, mX(6), 6)
    ccc = ccc + ((ddd Or (Not eee)) Xor aaa) + mX(6) + &H6D703EF3
    ccc = ((ccc * &H40&) Or (((ccc And &HFC000000) \ &H4000000) And &H3F&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(bbb, ccc, ddd, eee, aaa, mX(9), 14)
    bbb = bbb + ((ccc Or (Not ddd)) Xor eee) + mX(9) + &H6D703EF3
    bbb = ((bbb * &H4000&) Or (((bbb And &HFFFC0000) \ &H40000) And &H3FFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(aaa, bbb, ccc, ddd, eee, mX(11), 12)
    aaa = aaa + ((bbb Or (Not ccc)) Xor ddd) + mX(11) + &H6D703EF3
    aaa = ((aaa * &H1000&) Or (((aaa And &HFFF00000) \ &H100000) And &HFFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(eee, aaa, bbb, ccc, ddd, mX(8), 13)
    eee = eee + ((aaa Or (Not bbb)) Xor ccc) + mX(8) + &H6D703EF3
    eee = ((eee * &H2000&) Or (((eee And &HFFF80000) \ &H80000) And &H1FFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(ddd, eee, aaa, bbb, ccc, mX(12), 5)
    ddd = ddd + ((eee Or (Not aaa)) Xor bbb) + mX(12) + &H6D703EF3
    ddd = ((ddd * &H20&) Or (((ddd And &HF8000000) \ &H8000000) And &H1F&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(ccc, ddd, eee, aaa, bbb, mX(2), 14)
    ccc = ccc + ((ddd Or (Not eee)) Xor aaa) + mX(2) + &H6D703EF3
    ccc = ((ccc * &H4000&) Or (((ccc And &HFFFC0000) \ &H40000) And &H3FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(bbb, ccc, ddd, eee, aaa, mX(10), 13)
    bbb = bbb + ((ccc Or (Not ddd)) Xor eee) + mX(10) + &H6D703EF3
    bbb = ((bbb * &H2000&) Or (((bbb And &HFFF80000) \ &H80000) And &H1FFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(aaa, bbb, ccc, ddd, eee, mX(0), 13)
    aaa = aaa + ((bbb Or (Not ccc)) Xor ddd) + mX(0) + &H6D703EF3
    aaa = ((aaa * &H2000&) Or (((aaa And &HFFF80000) \ &H80000) And &H1FFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnHHH(eee, aaa, bbb, ccc, ddd, mX(4), 7)
    eee = eee + ((aaa Or (Not bbb)) Xor ccc) + mX(4) + &H6D703EF3
    eee = ((eee * &H80&) Or (((eee And &HFE000000) \ &H2000000) And &H7F&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnHHH(ddd, eee, aaa, bbb, ccc, mX(13), 5)
    ddd = ddd + ((eee Or (Not aaa)) Xor bbb) + mX(13) + &H6D703EF3
    ddd = ((ddd * &H20&) Or (((ddd And &HF8000000) \ &H8000000) And &H1F&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGGG(ccc, ddd, eee, aaa, bbb, mX(8), 15)
    ccc = ccc + ((ddd And eee) Or ((Not ddd) And aaa)) + mX(8) + &H7A6D76E9
    ccc = ((ccc * &H8000&) Or (((ccc And &HFFFE0000) \ &H20000) And &H7FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGGG(bbb, ccc, ddd, eee, aaa, mX(6), 5)
    bbb = bbb + ((ccc And ddd) Or ((Not ccc) And eee)) + mX(6) + &H7A6D76E9
    bbb = ((bbb * &H20&) Or (((bbb And &HF8000000) \ &H8000000) And &H1F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(aaa, bbb, ccc, ddd, eee, mX(4), 8)
    aaa = aaa + ((bbb And ccc) Or ((Not bbb) And ddd)) + mX(4) + &H7A6D76E9
    aaa = ((aaa * &H100&) Or (((aaa And &HFF000000) \ &H1000000) And &HFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(eee, aaa, bbb, ccc, ddd, mX(1), 11)
    eee = eee + ((aaa And bbb) Or ((Not aaa) And ccc)) + mX(1) + &H7A6D76E9
    eee = ((eee * &H800&) Or (((eee And &HFFE00000) \ &H200000) And &H7FF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(ddd, eee, aaa, bbb, ccc, mX(3), 14)
    ddd = ddd + ((eee And aaa) Or ((Not eee) And bbb)) + mX(3) + &H7A6D76E9
    ddd = ((ddd * &H4000&) Or (((ddd And &HFFFC0000) \ &H40000) And &H3FFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(ccc, ddd, eee, aaa, bbb, mX(11), 14)
    ccc = ccc + ((ddd And eee) Or ((Not ddd) And aaa)) + mX(11) + &H7A6D76E9
    ccc = ((ccc * &H4000&) Or (((ccc And &HFFFC0000) \ &H40000) And &H3FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGGG(bbb, ccc, ddd, eee, aaa, mX(15), 6)
    bbb = bbb + ((ccc And ddd) Or ((Not ccc) And eee)) + mX(15) + &H7A6D76E9
    bbb = ((bbb * &H40&) Or (((bbb And &HFC000000) \ &H4000000) And &H3F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(aaa, bbb, ccc, ddd, eee, mX(0), 14)
    aaa = aaa + ((bbb And ccc) Or ((Not bbb) And ddd)) + mX(0) + &H7A6D76E9
    aaa = ((aaa * &H4000&) Or (((aaa And &HFFFC0000) \ &H40000) And &H3FFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(eee, aaa, bbb, ccc, ddd, mX(5), 6)
    eee = eee + ((aaa And bbb) Or ((Not aaa) And ccc)) + mX(5) + &H7A6D76E9
    eee = ((eee * &H40&) Or (((eee And &HFC000000) \ &H4000000) And &H3F&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGGG(ddd, eee, aaa, bbb, ccc, mX(12), 9)
    ddd = ddd + ((eee And aaa) Or ((Not eee) And bbb)) + mX(12) + &H7A6D76E9
    ddd = ((ddd * &H200&) Or (((ddd And &HFF800000) \ &H800000) And &H1FF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(ccc, ddd, eee, aaa, bbb, mX(2), 12)
    ccc = ccc + ((ddd And eee) Or ((Not ddd) And aaa)) + mX(2) + &H7A6D76E9
    ccc = ((ccc * &H1000&) Or (((ccc And &HFFF00000) \ &H100000) And &HFFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(bbb, ccc, ddd, eee, aaa, mX(13), 9)
    bbb = bbb + ((ccc And ddd) Or ((Not ccc) And eee)) + mX(13) + &H7A6D76E9
    bbb = ((bbb * &H200&) Or (((bbb And &HFF800000) \ &H800000) And &H1FF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(aaa, bbb, ccc, ddd, eee, mX(9), 12)
    aaa = aaa + ((bbb And ccc) Or ((Not bbb) And ddd)) + mX(9) + &H7A6D76E9
    aaa = ((aaa * &H1000&) Or (((aaa And &HFFF00000) \ &H100000) And &HFFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(eee, aaa, bbb, ccc, ddd, mX(7), 5)
    eee = eee + ((aaa And bbb) Or ((Not aaa) And ccc)) + mX(7) + &H7A6D76E9
    eee = ((eee * &H20&) Or (((eee And &HF8000000) \ &H8000000) And &H1F&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnGGG(ddd, eee, aaa, bbb, ccc, mX(10), 15)
    ddd = ddd + ((eee And aaa) Or ((Not eee) And bbb)) + mX(10) + &H7A6D76E9
    ddd = ((ddd * &H8000&) Or (((ddd And &HFFFE0000) \ &H20000) And &H7FFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnGGG(ccc, ddd, eee, aaa, bbb, mX(14), 8)
    ccc = ccc + ((ddd And eee) Or ((Not ddd) And aaa)) + mX(14) + &H7A6D76E9
    ccc = ((ccc * &H100&) Or (((ccc And &HFF000000) \ &H1000000) And &HFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(bbb, ccc, ddd, eee, aaa, mX(12), 8)
    bbb = bbb + (ccc Xor ddd Xor eee) + mX(12)
    bbb = ((bbb * &H100&) Or (((bbb And &HFF000000) \ &H1000000) And &HFF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFFF(aaa, bbb, ccc, ddd, eee, mX(15), 5)
    aaa = aaa + (bbb Xor ccc Xor ddd) + mX(15)
    aaa = ((aaa * &H20&) Or (((aaa And &HF8000000) \ &H8000000) And &H1F&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFFF(eee, aaa, bbb, ccc, ddd, mX(10), 12)
    eee = eee + (aaa Xor bbb Xor ccc) + mX(10)
    eee = ((eee * &H1000&) Or (((eee And &HFFF00000) \ &H100000) And &HFFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(ddd, eee, aaa, bbb, ccc, mX(4), 9)
    ddd = ddd + (eee Xor aaa Xor bbb) + mX(4)
    ddd = ((ddd * &H200&) Or (((ddd And &HFF800000) \ &H800000) And &H1FF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFFF(ccc, ddd, eee, aaa, bbb, mX(1), 12)
    ccc = ccc + (ddd Xor eee Xor aaa) + mX(1)
    ccc = ((ccc * &H1000&) Or (((ccc And &HFFF00000) \ &H100000) And &HFFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(bbb, ccc, ddd, eee, aaa, mX(5), 5)
    bbb = bbb + (ccc Xor ddd Xor eee) + mX(5)
    bbb = ((bbb * &H20&) Or (((bbb And &HF8000000) \ &H8000000) And &H1F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))
    
    'fnFFF(aaa, bbb, ccc, ddd, eee, mX(8), 14)
    aaa = aaa + (bbb Xor ccc Xor ddd) + mX(8)
    aaa = ((aaa * &H4000&) Or (((aaa And &HFFFC0000) \ &H40000) And &H3FFF&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(eee, aaa, bbb, ccc, ddd, mX(7), 6)
    eee = eee + (aaa Xor bbb Xor ccc) + mX(7)
    eee = ((eee * &H40&) Or (((eee And &HFC000000) \ &H4000000) And &H3F&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(ddd, eee, aaa, bbb, ccc, mX(6), 8)
    ddd = ddd + (eee Xor aaa Xor bbb) + mX(6)
    ddd = ((ddd * &H100&) Or (((ddd And &HFF000000) \ &H1000000) And &HFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(ccc, ddd, eee, aaa, bbb, mX(2), 13)
    ccc = ccc + (ddd Xor eee Xor aaa) + mX(2)
    ccc = ((ccc * &H2000&) Or (((ccc And &HFFF80000) \ &H80000) And &H1FFF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(bbb, ccc, ddd, eee, aaa, mX(13), 6)
    bbb = bbb + (ccc Xor ddd Xor eee) + mX(13)
    bbb = ((bbb * &H40&) Or (((bbb And &HFC000000) \ &H4000000) And &H3F&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(aaa, bbb, ccc, ddd, eee, mX(14), 5)
    aaa = aaa + (bbb Xor ccc Xor ddd) + mX(14)
    aaa = ((aaa * &H20&) Or (((aaa And &HF8000000) \ &H8000000) And &H1F&)) + eee
    ccc = ((ccc * &H400&) Or (((ccc And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(eee, aaa, bbb, ccc, ddd, mX(0), 15)
    eee = eee + (aaa Xor bbb Xor ccc) + mX(0)
    eee = ((eee * &H8000&) Or (((eee And &HFFFE0000) \ &H20000) And &H7FFF&)) + ddd
    bbb = ((bbb * &H400&) Or (((bbb And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(ddd, eee, aaa, bbb, ccc, mX(3), 13)
    ddd = ddd + (eee Xor aaa Xor bbb) + mX(3)
    ddd = ((ddd * &H2000&) Or (((ddd And &HFFF80000) \ &H80000) And &H1FFF&)) + ccc
    aaa = ((aaa * &H400&) Or (((aaa And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(ccc, ddd, eee, aaa, bbb, mX(9), 11)
    ccc = ccc + (ddd Xor eee Xor aaa) + mX(9)
    ccc = ((ccc * &H800&) Or (((ccc And &HFFE00000) \ &H200000) And &H7FF&)) + bbb
    eee = ((eee * &H400&) Or (((eee And &HFFC00000) \ &H400000) And &H3FF&))

    'fnFFF(bbb, ccc, ddd, eee, aaa, mX(11), 11)
    bbb = bbb + (ccc Xor ddd Xor eee) + mX(11)
    bbb = ((bbb * &H800&) Or (((bbb And &HFFE00000) \ &H200000) And &H7FF&)) + aaa
    ddd = ((ddd * &H400&) Or (((ddd And &HFFC00000) \ &H400000) And &H3FF&))

    ddd = ddd + cc + mState(1)
    mState(1) = mState(2) + dd + eee
    mState(2) = mState(3) + ee + aaa
    mState(3) = mState(4) + aa + bbb
    mState(4) = mState(0) + bb + ccc
    mState(0) = ddd
End Sub
#Else
Private Sub Transform(ByRef Bytes() As Byte, ByVal Index As Long)
    CopyMemory mX(0), Bytes(Index), 64
    
    Dim aa As Long
    Dim bb As Long
    Dim dd As Long
    Dim cc As Long
    Dim ee As Long
    Dim aaa As Long
    Dim bbb As Long
    Dim ccc As Long
    Dim ddd As Long
    Dim eee As Long
    
    aa = mState(0)
    bb = mState(1)
    cc = mState(2)
    dd = mState(3)
    ee = mState(4)
    aaa = aa
    bbb = bb
    ccc = cc
    ddd = dd
    eee = ee
    
    fnFF aa, bb, cc, dd, ee, mX(0), 11
    fnFF ee, aa, bb, cc, dd, mX(1), 14
    fnFF dd, ee, aa, bb, cc, mX(2), 15
    fnFF cc, dd, ee, aa, bb, mX(3), 12
    fnFF bb, cc, dd, ee, aa, mX(4), 5
    fnFF aa, bb, cc, dd, ee, mX(5), 8
    fnFF ee, aa, bb, cc, dd, mX(6), 7
    fnFF dd, ee, aa, bb, cc, mX(7), 9
    fnFF cc, dd, ee, aa, bb, mX(8), 11
    fnFF bb, cc, dd, ee, aa, mX(9), 13
    fnFF aa, bb, cc, dd, ee, mX(10), 14
    fnFF ee, aa, bb, cc, dd, mX(11), 15
    fnFF dd, ee, aa, bb, cc, mX(12), 6
    fnFF cc, dd, ee, aa, bb, mX(13), 7
    fnFF bb, cc, dd, ee, aa, mX(14), 9
    fnFF aa, bb, cc, dd, ee, mX(15), 8

    fnGG ee, aa, bb, cc, dd, mX(7), 7
    fnGG dd, ee, aa, bb, cc, mX(4), 6
    fnGG cc, dd, ee, aa, bb, mX(13), 8
    fnGG bb, cc, dd, ee, aa, mX(1), 13
    fnGG aa, bb, cc, dd, ee, mX(10), 11
    fnGG ee, aa, bb, cc, dd, mX(6), 9
    fnGG dd, ee, aa, bb, cc, mX(15), 7
    fnGG cc, dd, ee, aa, bb, mX(3), 15
    fnGG bb, cc, dd, ee, aa, mX(12), 7
    fnGG aa, bb, cc, dd, ee, mX(0), 12
    fnGG ee, aa, bb, cc, dd, mX(9), 15
    fnGG dd, ee, aa, bb, cc, mX(5), 9
    fnGG cc, dd, ee, aa, bb, mX(2), 11
    fnGG bb, cc, dd, ee, aa, mX(14), 7
    fnGG aa, bb, cc, dd, ee, mX(11), 13
    fnGG ee, aa, bb, cc, dd, mX(8), 12

    fnHH dd, ee, aa, bb, cc, mX(3), 11
    fnHH cc, dd, ee, aa, bb, mX(10), 13
    fnHH bb, cc, dd, ee, aa, mX(14), 6
    fnHH aa, bb, cc, dd, ee, mX(4), 7
    fnHH ee, aa, bb, cc, dd, mX(9), 14
    fnHH dd, ee, aa, bb, cc, mX(15), 9
    fnHH cc, dd, ee, aa, bb, mX(8), 13
    fnHH bb, cc, dd, ee, aa, mX(1), 15
    fnHH aa, bb, cc, dd, ee, mX(2), 14
    fnHH ee, aa, bb, cc, dd, mX(7), 8
    fnHH dd, ee, aa, bb, cc, mX(0), 13
    fnHH cc, dd, ee, aa, bb, mX(6), 6
    fnHH bb, cc, dd, ee, aa, mX(13), 5
    fnHH aa, bb, cc, dd, ee, mX(11), 12
    fnHH ee, aa, bb, cc, dd, mX(5), 7
    fnHH dd, ee, aa, bb, cc, mX(12), 5
    
    fnII cc, dd, ee, aa, bb, mX(1), 11
    fnII bb, cc, dd, ee, aa, mX(9), 12
    fnII aa, bb, cc, dd, ee, mX(11), 14
    fnII ee, aa, bb, cc, dd, mX(10), 15
    fnII dd, ee, aa, bb, cc, mX(0), 14
    fnII cc, dd, ee, aa, bb, mX(8), 15
    fnII bb, cc, dd, ee, aa, mX(12), 9
    fnII aa, bb, cc, dd, ee, mX(4), 8
    fnII ee, aa, bb, cc, dd, mX(13), 9
    fnII dd, ee, aa, bb, cc, mX(3), 14
    fnII cc, dd, ee, aa, bb, mX(7), 5
    fnII bb, cc, dd, ee, aa, mX(15), 6
    fnII aa, bb, cc, dd, ee, mX(14), 8
    fnII ee, aa, bb, cc, dd, mX(5), 6
    fnII dd, ee, aa, bb, cc, mX(6), 5
    fnII cc, dd, ee, aa, bb, mX(2), 12

    fnJJ bb, cc, dd, ee, aa, mX(4), 9
    fnJJ aa, bb, cc, dd, ee, mX(0), 15
    fnJJ ee, aa, bb, cc, dd, mX(5), 5
    fnJJ dd, ee, aa, bb, cc, mX(9), 11
    fnJJ cc, dd, ee, aa, bb, mX(7), 6
    fnJJ bb, cc, dd, ee, aa, mX(12), 8
    fnJJ aa, bb, cc, dd, ee, mX(2), 13
    fnJJ ee, aa, bb, cc, dd, mX(10), 12
    fnJJ dd, ee, aa, bb, cc, mX(14), 5
    fnJJ cc, dd, ee, aa, bb, mX(1), 12
    fnJJ bb, cc, dd, ee, aa, mX(3), 13
    fnJJ aa, bb, cc, dd, ee, mX(8), 14
    fnJJ ee, aa, bb, cc, dd, mX(11), 11
    fnJJ dd, ee, aa, bb, cc, mX(6), 8
    fnJJ cc, dd, ee, aa, bb, mX(15), 5
    fnJJ bb, cc, dd, ee, aa, mX(13), 6

    fnJJJ aaa, bbb, ccc, ddd, eee, mX(5), 8
    fnJJJ eee, aaa, bbb, ccc, ddd, mX(14), 9
    fnJJJ ddd, eee, aaa, bbb, ccc, mX(7), 9
    fnJJJ ccc, ddd, eee, aaa, bbb, mX(0), 11
    fnJJJ bbb, ccc, ddd, eee, aaa, mX(9), 13
    fnJJJ aaa, bbb, ccc, ddd, eee, mX(2), 15
    fnJJJ eee, aaa, bbb, ccc, ddd, mX(11), 15
    fnJJJ ddd, eee, aaa, bbb, ccc, mX(4), 5
    fnJJJ ccc, ddd, eee, aaa, bbb, mX(13), 7
    fnJJJ bbb, ccc, ddd, eee, aaa, mX(6), 7
    fnJJJ aaa, bbb, ccc, ddd, eee, mX(15), 8
    fnJJJ eee, aaa, bbb, ccc, ddd, mX(8), 11
    fnJJJ ddd, eee, aaa, bbb, ccc, mX(1), 14
    fnJJJ ccc, ddd, eee, aaa, bbb, mX(10), 14
    fnJJJ bbb, ccc, ddd, eee, aaa, mX(3), 12
    fnJJJ aaa, bbb, ccc, ddd, eee, mX(12), 6

    fnIII eee, aaa, bbb, ccc, ddd, mX(6), 9
    fnIII ddd, eee, aaa, bbb, ccc, mX(11), 13
    fnIII ccc, ddd, eee, aaa, bbb, mX(3), 15
    fnIII bbb, ccc, ddd, eee, aaa, mX(7), 7
    fnIII aaa, bbb, ccc, ddd, eee, mX(0), 12
    fnIII eee, aaa, bbb, ccc, ddd, mX(13), 8
    fnIII ddd, eee, aaa, bbb, ccc, mX(5), 9
    fnIII ccc, ddd, eee, aaa, bbb, mX(10), 11
    fnIII bbb, ccc, ddd, eee, aaa, mX(14), 7
    fnIII aaa, bbb, ccc, ddd, eee, mX(15), 7
    fnIII eee, aaa, bbb, ccc, ddd, mX(8), 12
    fnIII ddd, eee, aaa, bbb, ccc, mX(12), 7
    fnIII ccc, ddd, eee, aaa, bbb, mX(4), 6
    fnIII bbb, ccc, ddd, eee, aaa, mX(9), 15
    fnIII aaa, bbb, ccc, ddd, eee, mX(1), 13
    fnIII eee, aaa, bbb, ccc, ddd, mX(2), 11

    fnHHH ddd, eee, aaa, bbb, ccc, mX(15), 9
    fnHHH ccc, ddd, eee, aaa, bbb, mX(5), 7
    fnHHH bbb, ccc, ddd, eee, aaa, mX(1), 15
    fnHHH aaa, bbb, ccc, ddd, eee, mX(3), 11
    fnHHH eee, aaa, bbb, ccc, ddd, mX(7), 8
    fnHHH ddd, eee, aaa, bbb, ccc, mX(14), 6
    fnHHH ccc, ddd, eee, aaa, bbb, mX(6), 6
    fnHHH bbb, ccc, ddd, eee, aaa, mX(9), 14
    fnHHH aaa, bbb, ccc, ddd, eee, mX(11), 12
    fnHHH eee, aaa, bbb, ccc, ddd, mX(8), 13
    fnHHH ddd, eee, aaa, bbb, ccc, mX(12), 5
    fnHHH ccc, ddd, eee, aaa, bbb, mX(2), 14
    fnHHH bbb, ccc, ddd, eee, aaa, mX(10), 13
    fnHHH aaa, bbb, ccc, ddd, eee, mX(0), 13
    fnHHH eee, aaa, bbb, ccc, ddd, mX(4), 7
    fnHHH ddd, eee, aaa, bbb, ccc, mX(13), 5

    fnGGG ccc, ddd, eee, aaa, bbb, mX(8), 15
    fnGGG bbb, ccc, ddd, eee, aaa, mX(6), 5
    fnGGG aaa, bbb, ccc, ddd, eee, mX(4), 8
    fnGGG eee, aaa, bbb, ccc, ddd, mX(1), 11
    fnGGG ddd, eee, aaa, bbb, ccc, mX(3), 14
    fnGGG ccc, ddd, eee, aaa, bbb, mX(11), 14
    fnGGG bbb, ccc, ddd, eee, aaa, mX(15), 6
    fnGGG aaa, bbb, ccc, ddd, eee, mX(0), 14
    fnGGG eee, aaa, bbb, ccc, ddd, mX(5), 6
    fnGGG ddd, eee, aaa, bbb, ccc, mX(12), 9
    fnGGG ccc, ddd, eee, aaa, bbb, mX(2), 12
    fnGGG bbb, ccc, ddd, eee, aaa, mX(13), 9
    fnGGG aaa, bbb, ccc, ddd, eee, mX(9), 12
    fnGGG eee, aaa, bbb, ccc, ddd, mX(7), 5
    fnGGG ddd, eee, aaa, bbb, ccc, mX(10), 15
    fnGGG ccc, ddd, eee, aaa, bbb, mX(14), 8

    fnFFF bbb, ccc, ddd, eee, aaa, mX(12), 8
    fnFFF aaa, bbb, ccc, ddd, eee, mX(15), 5
    fnFFF eee, aaa, bbb, ccc, ddd, mX(10), 12
    fnFFF ddd, eee, aaa, bbb, ccc, mX(4), 9
    fnFFF ccc, ddd, eee, aaa, bbb, mX(1), 12
    fnFFF bbb, ccc, ddd, eee, aaa, mX(5), 5
    fnFFF aaa, bbb, ccc, ddd, eee, mX(8), 14
    fnFFF eee, aaa, bbb, ccc, ddd, mX(7), 6
    fnFFF ddd, eee, aaa, bbb, ccc, mX(6), 8
    fnFFF ccc, ddd, eee, aaa, bbb, mX(2), 13
    fnFFF bbb, ccc, ddd, eee, aaa, mX(13), 6
    fnFFF aaa, bbb, ccc, ddd, eee, mX(14), 5
    fnFFF eee, aaa, bbb, ccc, ddd, mX(0), 15
    fnFFF ddd, eee, aaa, bbb, ccc, mX(3), 13
    fnFFF ccc, ddd, eee, aaa, bbb, mX(9), 11
    fnFFF bbb, ccc, ddd, eee, aaa, mX(11), 11

    ddd = Helper.UAdd(Helper.UAdd(ddd, cc), mState(1))
    mState(1) = Helper.UAdd(Helper.UAdd(mState(2), dd), eee)
    mState(2) = Helper.UAdd(Helper.UAdd(mState(3), ee), aaa)
    mState(3) = Helper.UAdd(Helper.UAdd(mState(4), aa), bbb)
    mState(4) = Helper.UAdd(Helper.UAdd(mState(0), bb), ccc)
    mState(0) = ddd
End Sub

Private Function fnF(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    fnF = X Xor Y Xor z
End Function

Private Function fnG(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    fnG = (X And Y) Or ((Not X) And z)
End Function

Private Function fnH(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    fnH = (X Or (Not Y)) Xor z
End Function

Private Function fnI(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    fnI = (X And z) Or (Y And (Not z))
End Function

Private Function fnJ(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
    fnJ = X Xor (Y Or (Not z))
End Function

Private Sub fnFF(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(a, fnF(b, c, d)), X)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnGG(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnG(b, c, d)), X), &H5A827999)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnHH(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnH(b, c, d)), X), &H6ED9EBA1)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnII(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnI(b, c, d)), X), &H8F1BBCDC)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnJJ(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnJ(b, c, d)), X), &HA953FD4E)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnFFF(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(a, fnF(b, c, d)), X)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnGGG(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnG(b, c, d)), X), &H7A6D76E9)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnHHH(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnH(b, c, d)), X), &H6D703EF3)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnIII(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnI(b, c, d)), X), &H5C4DD124)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub

Private Sub fnJJJ(ByRef a As Long, ByVal b As Long, ByRef c As Long, ByVal d As Long, ByVal e As Long, ByVal X As Long, ByVal s As Long)
    a = Helper.UAdd(Helper.UAdd(Helper.UAdd(a, fnJ(b, c, d)), X), &H50A28BE6)
    a = Helper.UAdd(LRotate(a, s), e)
    c = LRotate(c, 10)
End Sub
#End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
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
'   RIPEMD160
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get RIPEMD160_CanReuseTransform() As Boolean
    RIPEMD160_CanReuseTransform = CanReuseTransform
End Property

Private Property Get RIPEMD160_CanTransformMultipleBlocks() As Boolean
    RIPEMD160_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub RIPEMD160_Clear()
    Clear
End Sub

Private Function RIPEMD160_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    RIPEMD160_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function RIPEMD160_Equals(Value As Variant) As Boolean
    RIPEMD160_Equals = Equals(Value)
End Function

Private Function RIPEMD160_GetHashCode() As Long
    RIPEMD160_GetHashCode = GetHashCode
End Function

Private Property Get RIPEMD160_Hash() As Byte()
    RIPEMD160_Hash = Hash
End Property

Private Property Get RIPEMD160_HashSize() As Long
    RIPEMD160_HashSize = HashSize
End Property

Private Sub RIPEMD160_Initialize()
    Initialize
End Sub

Private Property Get RIPEMD160_InputBlockSize() As Long
    RIPEMD160_InputBlockSize = InputBlockSize
End Property

Private Property Get RIPEMD160_OutputBlockSize() As Long
    RIPEMD160_OutputBlockSize = OutputBlockSize
End Property

Private Function RIPEMD160_ToString() As String
    RIPEMD160_ToString = ToString
End Function

Private Function RIPEMD160_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    RIPEMD160_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function RIPEMD160_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    RIPEMD160_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HMACSHA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: HMACSHA1
'

''
' Represents an SHA-1 implementation of the HMAC algorithm.
'
' @remarks Information about HMAC can be found at http://csrc.nist.gov/publications/fips/fips198/fips-198a.pdf
' <p>A Hash-based Message Authentication Code can be used to determine if messages sent over insecure channels
' have been modified. This works by each party at boths ends sharing a secret key used to sign the sent message and
' to also authenticate the received message. If the key were compromised, then the authentication cannot be guaranteed.</p>
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements KeyedHashAlgorithm
Implements HMAC
Implements ICryptoTransform

Private mBase As HMACBase


''
' Returns the name of the hash.
'
' @return The name of the hash.
'
Public Property Get HashName() As String
    HashName = mBase.HashName
End Property

''
' Sets the name of the hash.
'
' @param Value The name of the hash.
'
Public Property Let HashName(ByVal Value As String)
    mBase.HashName = Value
End Property

''
' Returns the key being used by the hash.
'
' @return The key being used.
'
Public Property Get Key() As Byte()
    Key = mBase.GetKey
End Property

''
' Sets the key to be used by the hash.
'
' @param Value The new key value.
'
Public Property Let Key(ByRef Value() As Byte)
    mBase.SetKey Value
End Property

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
' <p>If a null array is passed in for the OutputBuffer, then nothing is copied, however, the hash is still computed
' on the InputBuffer data.</p>
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    TransformBlock = mBase.TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

''
' Processes the final data and retrieves the hash value.
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
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = Cor.NewHMACBase(New SHA1CryptoServiceProvider, 64, "SHA1")
End Sub

Friend Sub Init(ByRef Key() As Byte, ByVal UseManagedSha1 As Boolean)
    If UseManagedSha1 Then
        Set mBase = Cor.NewHMACBase(New SHA1Managed, 64, "SHA1", Key)
    Else
        mBase.SetKey Key
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   HMAC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HMAC_CanReuseTransform() As Boolean
    HMAC_CanReuseTransform = CanReuseTransform
End Property

Private Property Get HMAC_CanTransformMultipleBlocks() As Boolean
    HMAC_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub HMAC_Clear()
    Clear
End Sub

Private Function HMAC_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    HMAC_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function HMAC_Equals(Value As Variant) As Boolean
    HMAC_Equals = Equals(Value)
End Function

Private Function HMAC_GetHashCode() As Long
    HMAC_GetHashCode = GetHashCode
End Function

Private Property Get HMAC_Hash() As Byte()
    HMAC_Hash = Hash
End Property

Private Property Let HMAC_HashName(ByRef Value As String)
    HashName = Value
End Property

Private Property Get HMAC_HashName() As String
    HMAC_HashName = HashName
End Property

Private Property Get HMAC_HashSize() As Long
    HMAC_HashSize = HashSize
End Property

Private Sub HMAC_Initialize()
    Initialize
End Sub

Private Property Get HMAC_InputBlockSize() As Long
    HMAC_InputBlockSize = InputBlockSize
End Property

Private Property Let HMAC_Key(Value() As Byte)
    mBase.SetKey Value
End Property

Private Property Get HMAC_Key() As Byte()
    HMAC_Key = Key
End Property

Private Property Get HMAC_OutputBlockSize() As Long
    HMAC_OutputBlockSize = OutputBlockSize
End Property

Private Function HMAC_ToString() As String
    HMAC_ToString = ToString
End Function

Private Function HMAC_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    HMAC_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function HMAC_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    HMAC_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
'   KeyedHashAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let KeyedHashAlgorithm_Key(Value() As Byte)
    mBase.SetKey Value
End Property

Private Property Get KeyedHashAlgorithm_Key() As Byte()
    KeyedHashAlgorithm_Key = mBase.GetKey
End Property

Private Property Get KeyedHashAlgorithm_CanReuseTransform() As Boolean
    KeyedHashAlgorithm_CanReuseTransform = CanReuseTransform
End Property

Private Property Get KeyedHashAlgorithm_CanTransformMultipleBlocks() As Boolean
    KeyedHashAlgorithm_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub KeyedHashAlgorithm_Clear()
    Clear
End Sub

Private Function KeyedHashAlgorithm_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    KeyedHashAlgorithm_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function KeyedHashAlgorithm_Equals(Value As Variant) As Boolean
    KeyedHashAlgorithm_Equals = Equals(Value)
End Function

Private Function KeyedHashAlgorithm_GetHashCode() As Long
    KeyedHashAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get KeyedHashAlgorithm_Hash() As Byte()
    KeyedHashAlgorithm_Hash = Hash
End Property

Private Property Get KeyedHashAlgorithm_HashSize() As Long
    KeyedHashAlgorithm_HashSize = HashSize
End Property

Private Sub KeyedHashAlgorithm_Initialize()
    Initialize
End Sub

Private Property Get KeyedHashAlgorithm_InputBlockSize() As Long
    KeyedHashAlgorithm_InputBlockSize = InputBlockSize
End Property

Private Property Get KeyedHashAlgorithm_OutputBlockSize() As Long
    KeyedHashAlgorithm_OutputBlockSize = OutputBlockSize
End Property

Private Function KeyedHashAlgorithm_ToString() As String
    KeyedHashAlgorithm_ToString = ToString
End Function

Private Function KeyedHashAlgorithm_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    KeyedHashAlgorithm_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function KeyedHashAlgorithm_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    KeyedHashAlgorithm_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


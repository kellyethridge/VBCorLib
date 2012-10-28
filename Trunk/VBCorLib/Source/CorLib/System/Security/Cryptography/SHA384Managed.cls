VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA384Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2006 Kelly Ethridge
'
'    This file is part of VBCorLib.
'
'    VBCorLib is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBCorLib is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: SHA384Managed
'

''
' Provides services to compute the SHA-384 hash value.
'
' @remarks The SHA-384 algorithm is implemented in Visual Basic.
' <p>Details can be found at http://csrc.nist.gov/CryptoToolkit/tkhash.html</p>
' @see HashAlgorithm
'
Option Explicit
Implements IObject
Implements HashAlgorithm
Implements SHA384
Implements ICryptoTransform

Private mBase As New SHA512Managed


''
' Returns if this service provider can be reused to transfrom more data.
'
' @return Always returns True.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = mBase.CanReuseTransform
End Property

''
' Returns if this transform can operate on multiple blocks at once.
'
' @return Always returns True.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = mBase.CanTransformMultipleBlocks
End Property

''
' Returns the computed hash.
'
' @return A 48-byte array containing the hash value.
' @remarks The TransformFinalBlock must be called before accessing
' this property or an InvalidOperationException will be thrown.
'
Public Property Get Hash() As Byte()
    Dim Ret() As Byte
    Ret = mBase.Hash
    ReDim Preserve Ret(0 To 47)
    Hash = Ret
End Property

''
' Returns the size of the computed hash.
'
' @return Always returns 384 bits.
'
Public Property Get HashSize() As Long
    HashSize = 384
End Property

''
' Returns the input block size this transform can work on.
'
' @return Always returns 1.
' @remarks When calling the TransformBlock method, the InputCount must be zero or a multiple of this value.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = mBase.InputBlockSize
End Property

''
' Returns the output block size this transform can produce.
'
' @return Always returns 1.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = mBase.OutputBlockSize
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
' @return A 48-byte hash value.
' @remarks If a Stream object is passed in, then the stream is continuously read and the hash calculated until
' there is no more data left to read from the stream.
'
Public Function ComputeHash(ByRef Source As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte()
    Dim Ret() As Byte
    Ret = mBase.ComputeHash(Source, Index, Count)
    ReDim Preserve Ret(0 To 47)
    ComputeHash = Ret
End Function

''
' Resets the hashing algorithm to start again.
'
Public Sub Initialize()
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
    Equals = Object.Equals(Me, Value)
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
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' @return Returns a string representation of this instance.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call mBase.Init(modStaticClasses.SHA384.StatePtr)
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
'   IObject Interface
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
'   SHA384 Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get SHA384_CanReuseTransform() As Boolean
    SHA384_CanReuseTransform = CanReuseTransform
End Property

Private Property Get SHA384_CanTransformMultipleBlocks() As Boolean
    SHA384_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Sub SHA384_Clear()
    Call Clear
End Sub

Private Function SHA384_ComputeHash(Source As Variant, Optional Index As Variant, Optional Count As Variant) As Byte()
    SHA384_ComputeHash = ComputeHash(Source, Index, Count)
End Function

Private Function SHA384_Equals(Value As Variant) As Boolean
    SHA384_Equals = Equals(Value)
End Function

Private Function SHA384_GetHashCode() As Long
    SHA384_GetHashCode = GetHashCode
End Function

Private Property Get SHA384_Hash() As Byte()
    SHA384_Hash = Hash
End Property

Private Property Get SHA384_HashSize() As Long
    SHA384_HashSize = HashSize
End Property

Private Sub SHA384_Initialize()
    Call Initialize
End Sub

Private Property Get SHA384_InputBlockSize() As Long
    SHA384_InputBlockSize = InputBlockSize
End Property

Private Property Get SHA384_OutputBlockSize() As Long
    SHA384_OutputBlockSize = OutputBlockSize
End Property

Private Function SHA384_ToString() As String
    SHA384_ToString = ToString
End Function

Private Function SHA384_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    SHA384_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function SHA384_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    SHA384_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function





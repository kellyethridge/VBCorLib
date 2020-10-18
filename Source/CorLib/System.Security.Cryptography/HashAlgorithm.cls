VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HashAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: HashAlgorithm
'

''
' Represents a common interface implemented by all hashing algorithms.
'
' @remarks A hash is a fixed-length data that can represent unique variable-length data.
' The size of the hash determines the ability for different data to have a unique hash value.
' Since hash values are fixed-length, there are a fixed number of possible hash values, which
' must represent an infinite number of variable-length data.
' <p>Hashes are powerful in that a tiny change in the input data will cause a large change in
' the computed hash value. This provides the ability to detect when data has been altered.</p>
' <p>It is possible for two variable-length data to have the same hash value, but the chances
' are so remote as to be practically impossible. The larger the hash value, the more impractical
' is becomes to find a duplicate hash value.</p>
'
' @see SHA1CryptoServiceProvider
' @see SHA1Managed
' @see SHA256Managed
' @see SHA384Managed
' @see SHA512Managed
' @see RIPEMD160Managed
' @see MD5CryptoServiceProvider
'
Option Explicit

''
' Returns if the hash algorithm implementation can be reused once a hash has been computed.
'
' @returns True if the implementation can be reused, False otherwise.
'
Public Property Get CanReuseTransform() As Boolean: End Property

''
' Returns if the hash algorithm implementation can processd multiple blocks of data at a time.
'
' @return True if multiple blocks of data can be processed at a time, False otherwise.
' @remarks The size of 1 block of data is defined by the InputBlockSize property.
'
Public Property Get CanTransformMultipleBlocks() As Boolean: End Property

''
' Returns the computed hash after all data has been processed.
'
' @return A byte array containing the computed hash.
' @remarks The size of the array will vary depending on the hash algorithm being performed.
'
Public Property Get Hash() As Byte(): End Property

''
' Returns the size of the hash value in bits.
'
' @return The size of the hash value in bits.
'
Public Property Get HashSize() As Long: End Property

''
' Returns the size of the input block size allowed.
'
' @returns The number of bytes a single input block is.
' @remarks When calling the TransformBlock method, the InputCount must equal this value.
' If the hash's CanTransformMultipleBlocks returns True, then the InputCount can be
' larger than the InputBlockSize, but must be an even multiple of InputBlockSize.
'
Public Property Get InputBlockSize() As Long: End Property

''
' Returns the size of the output block.
'
' @return The size of the output block.
' @remarks When calling the TransformBlock method, the return value will be 0 or a multiple of this value.
'
Public Property Get OutputBlockSize() As Long: End Property

''
' Releases any resources being held by the hash algorithm.
'
' @remarks Once this is called, the object is disposed and cannot be reused.
'
Public Sub Clear(): End Sub

''
' Computes the hash on a source in its entirety.
'
' @param InputStreamOrBuffer The source of data compute the hash for. This can be a byte array or object implementing the Stream interface.
' @param Index The starting index in a byte array source. If the source is not a byte array, this argument is ignored.
' @param Count The number of bytes to compute the hash on. If the source is not a byte array, this argument is ignored.
' @return The hash value computed from the source provided.
' @remarks If the source is a Stream based object, then the stream is read from and the hash continuously computed until
' the stream no longer returns data.
' <p>If data has already been processed using TransformBlock, then that data will be included in the hashing process.</p>
'
Public Function ComputeHash(ByRef InputStreamOrBuffer As Variant, Optional ByRef Index As Variant, Optional ByRef Count As Variant) As Byte(): End Function

''
' Returns if the Value is equal to this object instance.
'
' @param Value The value to compare against this object instance.
' @return True if they are considered equal, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long: End Function

''
' Resets a hashing object to begin hashing a new value.
'
' @remarks If Clear has been called, then the object cannot be re-initialized.
'
Public Sub Initialize(): End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String: End Function

''
' Continues to compute the hash value for blocks of data.
'
' @param InputBuffer The bytes to continue computing the hash from.
' @param InputOffset The index into the byte array to begin computing from.
' @param InputCount The number of bytes to be included in the hash computation.
' @param OutputBuffer The data after it has been processed. This will be the same as the input data, no changes are made.
' @param OutputOffset The starting index in the output buffer to place the processed data.
' @return The number of bytes that were processed.
' @remarks The OutputBuffer will contain the same plain text data as the input buffer. No transformation of the data
' is applied. The OutputBuffer array can be a Null array or the same array as the InputBuffer. If the InputBuffer is
' used as the OutputBuffer and the InputOffset equals the OutputOffset, no data is copied.
' <p>Once this method is called, the Hash property cannot be called until the TransformFinalBlock is called, finishing
' the hash computation.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long: End Function

''
' Finalizes the hash computation by processing the last block of data.
'
' @param InputBuffer The bytes to finish the hash computation with.
' @param InputOffset The index into the byte array to begin hash computations.
' @param InputCount The number of bytes to be included in the final hash computation.
' @return A copy of the portion of the InputBuffer that was processed.
' @remarks The hash value is not returned by this method. To retrieve the final
' hash value, call the Hash property.
' <p>The Hash property can only be called after this method is called. If the Hash property
' is called before this method, then an exception is thrown.</p>
' <p>If the implementation can be reused (CanReuseTransform) then the implemntation is reset
' to allow computing of a new hash value.</p>
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte(): End Function


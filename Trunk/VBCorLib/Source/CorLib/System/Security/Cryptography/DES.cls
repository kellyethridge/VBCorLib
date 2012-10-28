VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DES"
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
'    Module: DES
'

''
' Represents a base interface that all DES algorithms should implement.
'
' @remarks Data Encryption Standard (DES) is a block cipher that uses a 56bit key
' encryption mechanism.
'
' @see DESCryptoServiceProvider
' @see SymmetricAlgorithm
'
Option Explicit

' Inherited from SymmetricAlgorithm

''
' Returns the block size in bits.
'
' @return The block size in bits.
' @remarks This always returns 64.
' <p>The block size is the amount of data that is processed as a single unit
' by the cipher. It will always require a block of 64 bits (8 bytes) to operate,
' therefore if a the last block to be operated on is not 64 bits in length it
' must be padded to meet the block size requirement. The padding is automatically
' performed by the routine</p>
'
Public Property Get BlockSize() As Long: End Property

''
' Sets the block size in bits.
'
' @param RHS The size of the block in bits.
' @remarks This always returns 64.
' <p>The block size is the amount of data that is processed as a single unit
' by the cipher. It will always require a block of 64 bits (8 bytes) to operate,
' therefore if a the last block to be operated on is not 64 bits in length it
' must be padded to meet the block size requirement. The padding is automatically
' performed by the routine</p>
'
Public Property Let BlockSize(ByVal RHS As Long): End Property

''
' Returns the feedback size in bits.
'
' @return The feedback size in bits.
' @remarks The always returns 64 bits.
' <p>The feedback size determines how much of the previous data processed is to be used
' on the next block of data. This is to prevent identical blocks of data to appear the
' same when encrypted.</p>
'
Public Property Get FeedbackSize() As Long: End Property

''
' Sets the feedback size in bits.
'
' @param RHS The size of feedback in bits.
' @remarks DES only supports 64 bits.
' <p>The feedback size determines how much of the previous data processed is to be used
' on the next block of data. This is to prevent identical blocks of data to appear the
' same when encrypted.</p>
' <p>Not all modes of DES support feedback.</p>
'
Public Property Let FeedbackSize(ByVal RHS As Long): End Property

''
' Returns the Initialization Vector (IV).
'
' @return An array representing the IV.
' @remarks Only an 8-byte array is supported.
' <p>When the very first block of data is process there is no previous data to pull
' feedback data from, so the Initialization Vector is used as a dummy set of data for feedback.</p>
'
Public Property Get IV() As Byte(): End Property

''
' Sets the Initialization Vectory (IV).
'
' @param RHS A set of bytes representing the IV.
' @remarks Only an 8-byte array is supported.
' <p>When the very first block of data is process there is no previous data to pull
' feedback data from, so the Initialization Vector is used as a dummy set of data for feedback.</p>
'
Public Property Let IV(ByRef RHS() As Byte): End Property

''
' Returns the secret key used by the cryptographic operation.
'
' @return Returns the secret key.
' @remarks Valid key sizes are determined by the individual algorithms. A list of valid sizes can be obtained
' through the <b>LegalKeySizes</b> property.
'
Public Property Get Key() As Byte(): End Property

''
' Sets the secret key to be used by the cryptographic operation.
'
' @param RHS The new secret key to be used.
' @remarks Valid key sizes are determined by the individual algorithms. A list of valid sizes can be obtained
' through the <b>LegalKeySizes</b> property.
'
Public Property Let Key(ByRef RHS() As Byte): End Property

''
' Returns the size of the key, in bits, used by the cryptographic operation.
'
' @return The size of the key in bits.
' @remarks This will aways return 64.
'
Public Property Get KeySize() As Long: End Property

''
' Sets the size of the key, in bits.
'
' @param RHS The size of the key in bits.
' @remarks This will cause the <b>Key</b> property to reset and regenerate when needed.
'
Public Property Let KeySize(ByVal RHS As Long): End Property

''
' Returns a list of legal block sizes.
'
' @return A list of legal block sizes the cryptographic operation supports.
' @remarks Block sizes are represented in bits.
'
Public Property Get LegalBlockSizes() As KeySizes(): End Property

''
' Returns a list of legal key sizes.
'
' @return A list of legal key sizes the cryptographic operation supports.
' @remarks Key sizes are represented in bits.
'
Public Property Get LegalKeySizes() As KeySizes(): End Property
''
' Returns the mode the current algoritm is set to.
'
' @return Returns a <b>CipherMode</b> value.
' @remarks The algoritm can operate in several modes. The modes determine how the data
' is processed from one block to the next.
'
Public Property Get Mode() As CipherMode: End Property

''
' Sets the mode the current algorithm is set to.
'
' @param RHS The new mode for the algorithm.
' @remarks The algoritm can operate in several modes. The modes determine how the data
' is processed from one block to the next.
'
Public Property Let Mode(ByVal RHS As CipherMode): End Property

''
' Returns the padding to be used by the algorithm.
'
' @return A current padding being used.
' @remarks Padding is added to the end of plain text data to make the data length
' an even multiple of the block size.
'
Public Property Get Padding() As PaddingMode: End Property

''
' Sets the padding to be used by the algorithm.
'
' @param RHS The new padding value to be used.
' @remarks Padding is added to the end of plain text data to make the data length
' an even multiple of the block size.
'
Public Property Let Padding(ByVal RHS As PaddingMode): End Property

''
' Releases all resources.
'
Public Sub Clear(): End Sub

''
' Creates a new decryptor using the Data Encryption Standard algorithm.
'
' @param RgbKey The secret key to be used to decrypt the ciphered text back into plain text.
' @param RgbIV The Initialization Vector to begin the feedback with.
' @return A new <b>ICryptoTransform</b> used to decrypt ciphered text.
' @remarks Though both parameters are optional, if only one parameter is supplied, an exception will be thrown.
'
Public Function CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform: End Function

''
' Creates a new encryptor using the Data Ecryption Standard algorithm.
'
' @param RgbKey The secret key to be used to encrypt the plain text into ciphered text.
' @param RgbIV The Initialization Vector to begin the feedback with.
' @return A new <b>ICryptoTransform</b> used to encrypt plain text.
' @remarks Though both parameters are optional, if only one parameter is supplied, an exception will be thrown.
'
Public Function CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform: End Function

''
' Returns if this instance and the value are equal.
'
' @param Value The value to compare for equality.
' @return Returns if the value and this instance are equal.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Generates a new Initialization Vector array.
'
Public Sub GenerateIV(): End Sub

''
' Generates a new Key based on the key size.
'
Public Sub GenerateKey(): End Sub

''
' Returns a semi-unique value that represents this instance.
'
' @return Returns the semi-unique value.
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string that represents this instance.
'
Public Function ToString() As String: End Function

''
' Checks if the specific number of bits is a valid key size.
'
' @param BitLength A value to determine if is a legal key size.
' @return Returns True if the bit length is a valid key size, False otherwise.
'
Public Function ValidKeySize(ByVal BitLength As Long) As Boolean: End Function




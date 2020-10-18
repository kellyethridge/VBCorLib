VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: RC2
'

''
' Represents a standard interface implemented by all RC2 crypto service providers.
'
Option Explicit



''
' Returns the effective key size.
'
' @return The effective key size.
' @remarks The effective key size is the number of bits to actually be used
' from the key itself. This applies to Windows 2000.
'
Public Property Get EffectiveKeySize() As Long: End Property

''
' Sets the effective key size.
'
' @param Value The effective key size.
' @remarks The effective key size is the number of bits to actually be used
' from the key itself. This applies to Windows 2000.
'
Public Property Let EffectiveKeySize(ByVal Value As Long): End Property

''
' Returns the block size, in bits, used by the cryptographic operation.
'
' @Return The block size in bits.
' @remarks A block size is the basic unit size of data that can be encrypted or decrypted in a single operation. Data that is
' larger than a block size are processed by splitting the data into multiple blocks the same size as the block size then
' operating of the multiple blocks.
' <p>A list of legal block sizes can be obtained through the <b>LegalBlockSizes</b> property.</p>
'
Public Property Get BlockSize() As Long: End Property

''
' Sets the block size, in bits, used by the cryptographic operation.
'
' @param Value The block size in bits.
' @remarks A block size is the basic unit size of data that can be encrypted or decrypted in a single operation. Data that is
' larger than a block size are processed by splitting the data into multiple blocks the same size as the block size then
' operating of the multiple blocks.
' <p>An invalid block size will cause a <b>CryptographicException</b>.
' <p>A list of legal block sizes can be obtained through the <b>LegalBlockSizes</b> property.</p>
'
Public Property Let BlockSize(ByVal Value As Long): End Property

''
' Returns the feedback size, in bits, used by the cryptographic operation.
'
' @return Returns the feedback size in bits.
' @remarks The feedback size determines the amount of data that is fed back to successive encryption or decryption operations.
' The feedback size cannot be greater than the block size.
'
Public Property Get FeedbackSize() As Long: End Property

''
' Sets the feedback size, in bits, used by the cryptographic operation.
'
' @param Value The feedback size in bits.
' @remarks The feedback size determines the amount of data that is fed back to successive encryption or decryption operations.
' The feedback size cannot be greater than the block size.
'
Public Property Let FeedbackSize(ByVal Value As Long): End Property

''
' Returns the Initialization Vector used by the cryptographic operation.
'
' @return Returns an array of bytes representing an Initialization Vector.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same aswell. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using the <b>GenerateIV</b> method, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
'
Public Property Get IV() As Byte(): End Property

''
' Sets the Initialization Vector used by the cryptographic operation.
'
' @param Value The new Initialization Vector.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same aswell. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using the <b>GenerateIV</b> method, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
' <p>This must be set to an array the same size as the block size.</p>
'
Public Property Let IV(ByRef Value() As Byte): End Property

''
' Returns the secret key used by the cryptographic operation.
'
' @return Returns the secret key.
' @remarks Valid key sizes are determined by the individual algorithms. A list of valid sizes can be obtained
' through the <b>LegalKeySizes</b> property.
' <p>If a key has not been set, then a new key will be generated automatically.</p>
'
Public Property Get Key() As Byte(): End Property

''
' Sets the secret key to be used by the cryptographic operation.
'
' @param Value The new secret key to be used.
' @remarks Valid key sizes are determined by the individual algorithms. A list of valid sizes can be obtained
' through the <b>LegalKeySizes</b> property.
'
Public Property Let Key(ByRef Value() As Byte): End Property

''
' Returns the size of the key, in bits, used by the cryptographic operation.
'
' @return The size of the key in bits.
'
Public Property Get KeySize() As Long: End Property

''
' Sets the size of the key, in bits.
'
' @param Value The size of the key in bits.
' @remarks This will cause the <b>Key</b> property to reset and regenerate when needed.
'
Public Property Let KeySize(ByVal Value As Long): End Property

''
' Returns a list of legal block sizes.
'
' @return A list of legal block sizes the cryptographic operation supports.
'
Public Property Get LegalBlockSizes() As KeySizes(): End Property

''
' Returns a list of legal key sizes.
'
' @return A list of legal key sizes the cryptographic operation supports.
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
' @param Value The new mode for the algorithm.
' @remarks The algoritm can operate in several modes. The modes determine how the data
' is processed from one block to the next.
'
Public Property Let Mode(ByVal Value As CipherMode): End Property

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
' @param Value The new padding value to be used.
' @remarks Padding is added to the end of plain text data to make the data length
' an even multiple of the block size.
'
Public Property Let Padding(ByVal Value As PaddingMode): End Property

''
' Releases all resources.
'
Public Sub Clear(): End Sub

''
' Creates a new decryptor using the RC2 algorithm.
'
' @param RgbKey The secret key to be used to decrypt the ciphered text back into plain text.
' @param RgbIV The Initialization Vector to begin the feedback with.
' @return A new <b>ICryptoTransform</b> used to decrypt ciphered text.
' @remarks Though both parameters are optional, if only one parameter is supplied, an exception will be thrown.
'
Public Function CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform: End Function

''
' Creates a new encryptor using the RC2 algorithm.
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





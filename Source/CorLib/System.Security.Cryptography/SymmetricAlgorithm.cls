VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymmetricAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: SymmetricAlgorithm
'

''
' Provides a standard interface for utilizing symmetric algorithms.
'
' @remarks A symmetric algorithm is one that uses the same Key and Initialization Vector to both encrypt and decrypt data.
' <p>For one of the symmetric algorithms to work, a Key and Initialization Vector needs to be supplied
' to the Crypto Service Provider.</p>
' <p><a href="RijndaelManaged.html">RijndaelManaged</a>, <a href="DESCryptoServiceProvider.html">DESCryptoServiceProvider</a>,
' <a href="RC2CryptoServiceProvider.html">RC2CryptoServiceProvider</a>, and <a href="TripleDESCryptoServiceProvider.html">TripleDESCryptoServiceProvider</a> are implementations of symmetric algorithms.</p>
' @see SymmetricAlgorithmBase
Option Explicit

''
' A defined set of cipher modes the algorithms can use.
'
' @param CBC The Cipher Block Chaining (CBC) uses a feedback mechanism by Exclusive ORing the previously encrypted block of data
' with the current block being processed. This is to prevent two identical blocks from encrypting into the same to encrypted
' blocks of data, making it more difficult to break the encryption. An Initialization Vector array is also used by Exclusive
' ORing the vector with the current block before encryption.
' @param ECB The Electronic Codebook (ECB) processes blocks individually and does not alter identical blocks of data.
' If two or more blocks of data are the same then their encrypted blocks will also be the same. A message with several
' identical blocks could become easier to decrypt.
' @param OFB The Output Feedback (OFB) method processes small increments of blocks into ciphered text instead of the entire
' block at one time. This works similar to the Cipher Feedback method, except if a ciphered bit is corrupted, then all plain text
' data following the bit will be corrupted.
' @param CFB The Cipher Feedback (CFB) method processes small increments of plain text into ciphered text instead of the entire
' block at one time. The method uses a shift register that is one block in length. If one bit of ciphered data is corrupted, then
' the corrisponding plain text bit is also corrupted until the corrupted bit is finally shifted out of the register.
' @param CTS The Cipher Text Stealing (CTS) mode handles any length of plain text and produces cipher text whose length matches
' the plain text length. This mode behaves like the CBC mode for all but the last two blocks of the plain text.
Public Enum CipherMode
    CBC = 1
    ECB = 2
    OFB = 3
    CFB = 4
    CTS = 5
End Enum

''
' A defined set of padding modes used to fill the end of non-blocksize length data blocks. For example if an algorithm
' has a blocksize of 8 bytes and the data is 3 bytes in length, then 5 bytes of padding must be applied.
'
' @param None No padding is applied to the data.
' @param PKCS7 The padding that is used is defined by a pattern of filling the remaining bytes with the numeric value
' of the padding bytes needed. If 5 bytes of padding are needed, then the 5 bytes would be 05 05 05 05 05.
' @param Zeros The padded bytes are all equal to zero.
' @param ANSIX923 The padding that is used is defined by a pattern of bytes set to zero with the last byte having the
' value of the number of padded bytes added. If 5 bytes of padding are needed, then the 5 bytes would be 00 00 00 00 05.
' @param ISO10126 The padded bytes are random in value with the last byte being the number of padding bytes. If 5 bytes
' of padding are needed, then the 5 bytes might be 23 9A 48 7C 05.
' @remarks Most plain text messages do not have an length that is a multiple of the block size to be processed by the cipher.
' There is usually not enough bytes to completely fill the last block to be processed. In these cases, additional bytes need
' to be appended to make the final block the same size as the block size being processed.
' <p>The reverse is done when decrypting ciphered data back to plain text. The padded bytes are removed.</p>
' <p>If a plain text message is 9 bytes long and the block size to be processed is 8 bytes, then two blocks will need to be
' processed. However, the second block will only have 1 byte, so 7 additional bytes are added for padding.</p>
' <p>Data: FF FF FF FF FF FF FF FF FF</p>
' <p>None: FF FF FF FF FF FF FF FF FF</p>
' <p>PKCS7: FF FF FF FF FF FF FF FF FF 07 07 07 07 07 07 07</p>
' <p>ANSIX923: FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 07</p>
' <p>ISO10126: FF FF FF FF FF FF FF FF FF 7D 2A 75 EF F8 EF 07</p>
' <p>Zeros: FF FF FF FF FF FF FF FF FF 00 00 00 00 00 00 00</p>
'
Public Enum PaddingMode
    None = 1
    PKCS7 = 2
    Zeros = 3
    ANSIX923 = 4
    ISO10126 = 5
End Enum


''
' Gets the block size, in bits, used by the cryptographic operation.
'
' @Return The block size in bits.
' @remarks The block size is the basic unit of data that can be encrypted or decrypted in one operation.
' Messages longer than the block size are handled as successive blocks; messages shorter than the block size
' must be padded with extra bits to reach the size of a block. Valid block sizes are determined by the symmetric algorithm used.
' <p>A list of valid block sizes can be obtained through the <a href="SymmetricAlgorithm.Get.LegalBlockSizes.html">LegalBlockSizes</a> property.</p>
' @see #LegalBlockSizes
Public Property Get BlockSize() As Long: End Property

''
' Sets the block size, in bits, used by the cryptographic operation.
'
' @param Value The block size in bits.
' @remarks The block size is the basic unit of data that can be encrypted or decrypted in one operation.
' Messages longer than the block size are handled as successive blocks; messages shorter than the block size
' must be padded with extra bits to reach the size of a block. Valid block sizes are determined by the symmetric algorithm used.
' <p>A list of valid block sizes can be obtained through the <a href="SymmetricAlgorithm.Get.LegalBlockSizes.html">LegalBlockSizes</a> property.</p>
' @see #LegalBlockSizes
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>The block size is invalid.</td>
'   </tr>
' </table>
Public Property Let BlockSize(ByVal Value As Long): End Property

''
' Gets the feedback size, in bits, used by the cryptographic operation.
'
' @return The feedback size in bits.
' @remarks The feedback size determines the amount of data that is fed back to successive encryption or decryption operations.
' The feedback size cannot be greater than the block size.
Public Property Get FeedbackSize() As Long: End Property

''
' Sets the feedback size, in bits, used by the cryptographic operation.
'
' @param Value The feedback size in bits.
' @remarks The feedback size determines the amount of data that is fed back to successive encryption or decryption operations.
' The feedback size cannot be greater than the block size.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>The feedback size is larger than the block size.</td>
'   </tr>
' </table>
Public Property Let FeedbackSize(ByVal Value As Long): End Property

''
' Gets the initialization vector (<b>IV</b>) for the symmetric algorithm.
'
' @return The initialization vector.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same aswell. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using <a href="SymmetricAlgorithm.GenerateIV.html">GenerateIV</a>, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
' <p>This must be set to an array the same size as the block size divided by 8.</p>
Public Property Get IV() As Byte(): End Property

''
' Sets the initialization vector (<b>IV</b>) for the symmetric algorithm.
'
' @param Value The initialization vector.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same aswell. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using the <a href="SymmetricAlgorithm.GenerateIV.html">GenerateIV</a> method, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
' <p>This must be set to an array the same size as the block size divided by 8.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>An attempt was made to set the initialization vector to an unitialized array.</td>
'   </tr>
'    <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>An attempt was made to set the initialization vector to an invalid size.</td>
'    </tr>
' </table>
Public Property Let IV(ByRef Value() As Byte): End Property

''
' Gets the secret key for the symmetric algorithm.
'
' @return The secret key to use for the symmetric algorithm.
' @remarks The secret key is used both for encryption and for decryption. For a symmetric algorithm to be successful, the secret
' key must be known only to the sender and the receiver. The valid key sizes are specified by the particular symmetric algorithm
' implementation and are listed in the <a href="SymmetircAlgorithm.Get.LegalKeySizes.html">LegalKeySizes</a> property.
' <p>If a key has not been set, the <a href="SymmetricAlgorithm.GenerateKey.html">GenerateKey</a> method will be called to create a new random value.</p>
Public Property Get Key() As Byte(): End Property

''
' Gets the secret key for the symmetric algorithm.
'
' @param Value The secret key to use for the symmetric algorithm.
' @remarks The secret key is used both for encryption and for decryption. For a symmetric algorithm to be successful, the secret
' key must be known only to the sender and the receiver. The valid key sizes are specified by the particular symmetric algorithm
' implementation and are listed in the <a href="SymmetircAlgorithm.Get.LegalKeySizes.html">LegalKeySizes</a> property.
' <p>If a key has not been set, the <a href="SymmetricAlgorithm.GenerateKey.html">GenerateKey</a> method will be called to create a new random value.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>An attempt was made to set the key to an uninitialized array.</td>
'   </tr>
'    <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>The key size is invalid.</td>
'    </tr>
' </table>
Public Property Let Key(ByRef Value() As Byte): End Property

''
' Gets the size of the key, in bits, used by the symmetric algorithm.
'
' @return The size, in bits, of the secret key used by the symmetric algorithm.
' @remarks The valid key sizes are specified by the particular symmetric algorithm implementation and are listed in the <a href="SymmetircAlgorithm.Get.LegalKeySizes.html">LegalKeySizes</a> property.
Public Property Get KeySize() As Long: End Property

''
' Gets the size of the key, in bits, used by the symmetric algorithm.
'
' @param Value The size, in bits, of the secret key used by the symmetric algorithm.
' @remarks The valid key sizes are specified by the particular symmetric algorithm implementation and are listed in the <a href="SymmetircAlgorithm.Get.LegalKeySizes.html">LegalKeySizes</a> property.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>The key size is not valid.</td>
'   </tr>
' </table>
Public Property Let KeySize(ByVal Value As Long): End Property

''
' Gets the block sizes, in bits, that are supported by the symmetric algorithm.
'
' @return An array that contains the block sizes supported by the algorithm.
' @remarks The symmetric algorithm supports only block sizes that match an entry in this array.
Public Property Get LegalBlockSizes() As KeySizes(): End Property

''
' Gets the key sizes, in bits, that are supported by the symmetric algorithm.
'
' @return An array that contains the key sizes supported by the algorithm.
' @remarks The symmetric algorithm supports only key sizes that match an entry in this array.
Public Property Get LegalKeySizes() As KeySizes(): End Property

''
' Gets the mode for operation of the symmetric algorithm.
'
' @return The mode for operation of the symmetric algorithm. The default is <a href="SymmetricAlgorithm.CipherMode.html">CipherMode.CBC</a>.
' @remarks See <a href="SymmetricAlgorithm.CipherMode.html">CipherMode</a> enumeration for a description of specific modes.
Public Property Get Mode() As CipherMode: End Property

''
' Sets the mode for operation of the symmetric algorithm.
'
' @param Value The mode for operation of the symmetric algorithm. The default is <a href="SymmetricAlgorithm.CipherMode.html">CipherMode.CBC</a>.
' @remarks See <a href="SymmetricAlgorithm.CipherMode.html">CipherMode</a> enumeration for a description of specific modes.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="CryptographicException.html">CryptographicException</a></td><td>The cipher mode is not one of the <a href="SymmetricAlgorithm.CipherMode.html">CipherMode</a> values.</td>
'   </tr>
' </table>
Public Property Let Mode(ByVal Value As CipherMode): End Property

''
' Gets the padding mode used in the symmetric algorithm.
'
' @return The padding mode used in the symmetric algorithm. The default is <a href="SymmetricAlgorithm.PaddingMode.html">PaddingMode.PKCS7</a>.
' @remarks Most plain text messages do not consist of a number of bytes that completely fill blocks. Often, there are not enough bytes
' to fill the last block. When this happens, a padding string is added to the text. For example, if the block length is 64 bits and
' the last block contains only 40 bits, 24 bits of padding are added. See the <a href="SymmetricAlgorithm.PaddingMode.html">PaddingMode</a> enumeration for a description of specific modes.
Public Property Get Padding() As PaddingMode: End Property

''
' Gets the padding mode used in the symmetric algorithm.
'
' @param Value The padding mode used in the symmetric algorithm. The default is <a href="SymmetricAlgorithm.PaddingMode.html">PaddingMode.PKCS7</a>.
' @remarks Most plain text messages do not consist of a number of bytes that completely fill blocks. Often, there are not enough bytes
' to fill the last block. When this happens, a padding string is added to the text. For example, if the block length is 64 bits and
' the last block contains only 40 bits, 24 bits of padding are added. See the <a href="SymmetricAlgorithm.PaddingMode.html">PaddingMode</a> enumeration for a description of specific modes.
Public Property Let Padding(ByVal Value As PaddingMode): End Property

''
' Releases all resources used by the <a href="SymmetricAlgorithm.html">SymmetricAlgorithm</a> class.
'
Public Sub Clear(): End Sub

''
' When implemented by a class, creates a symmetric decryptor object with the specified
' <a href="SymmetricAlgorithm.Get.Key.html">Key</a> property and initialization vector (<a href="SymmetricAlgorithm.Get.IV.html">IV</a>).
'
' @param RgbKey The secret key to use for the symmetric algorithm.
' @param RgbIV The initialization vector to use for the symmetric algorithm.
' @return A symmetric decryptor object.
' @remarks This method decrypts an encrypted message created using the <a href="SymmetricAlgorithm.CreateEncryptor.html">CreateEncryptor</a> with the same parameters.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>Only <i>RgbKey</i> or <i>RgbIV</i> was specified.</td>
'   </tr>
' </table>
' @see #CreateEncryptor
Public Function CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform: End Function

''
' When implemented by a class, creates a symmetric encryptor object with the specified
' <a href="SymmetricAlgorithm.Get.Key.html">Key</a> property and initialization vector (<a href="SymmetricAlgorithm.Get.IV.html">IV</a>).
'
' @param RgbKey The secret key to use for the symmetric algorithm.
' @param RgbIV The initialization vector to use for the symmetric algorithm.
' @return A symmetric encryptor object.
' @remarks Use the <a href="SymmetricAlgorithm.CreateDecryptor.html">CreateDecryptor</a> overload with the same parameters to decrypt the result of this method.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>Only <i>RgbKey</i> or <i>RgbIV</i> was specified.</td>
'   </tr>
' </table>
' @see #CreateDecryptor
Public Function CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform: End Function

''
' Returns if this instance and the value are equal.
'
' @param Value The value to compare for equality.
' @return Returns if the value and this instance are equal.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' When implemented by a class, generates a random initialization vector (<a href="SymmetricAlgorithm.Get.IV.html">IV</a>) to use for the algorithm.
'
' @remarks In general, there is no reason to use this method, because <code>CreateEncryptor()</code> automatically generates both an initialization vector and a key.
' However, you may want to use the <b>GenerateIV()</b> method to reuse a symmetric algorithm instance with a different initialization vector.
Public Sub GenerateIV(): End Sub

''
' When implemented ba a class, generates a random key (<a href="SymmetricAlgorithm.Get.Key.html">Key</a>) to use for the algorithm.
'
' @remarks In general, there is no reason to use this method, because <code>CreateEncryptor()</code> automatically generates both an initialization vector and a key.
' However, you may want to use the <b>GenerateKey()</b> method to reuse a symmetric algorithm instance with a different key.
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
' Determines whether the specified key size is valid for the current algorithm.
'
' @param BitLength The length, in bits, to check for a valid key size.
' @return <b>True</b> if the specified key size is valid for the current algorithm; otherwise, <b>False</b>.
Public Function ValidKeySize(ByVal BitLength As Long) As Boolean: End Function


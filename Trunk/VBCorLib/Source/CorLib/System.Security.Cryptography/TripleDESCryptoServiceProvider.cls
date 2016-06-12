VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TripleDESCryptoServiceProvider"
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
' Module: TripleDESCryptoServiceProvider
'

''
' Provides cryptographic services using the TripleDES cryptographic service provider.
'
' @see TripleDESStatic
' @include "..\DocIncludes\TripleDESCryptoServiceProvider.html"
Option Explicit
Implements IObject
Implements SymmetricAlgorithm
Implements TripleDES

Private Const LegalBlockSize        As Long = 64
Private Const LegalMinKeySize       As Long = 128
Private Const LegalMaxKeySize       As Long = 192
Private Const DefaultKeySize        As Long = LegalMaxKeySize
Private Const LegalSkipSize         As Long = 64
Private Const DefaultFeedbackSize   As Long = 8

Private WithEvents mBase As SymmetricAlgorithmBase
Attribute mBase.VB_VarHelpID = -1



''
' Returns the block size, in bits, used by this algorithm.
'
' @return The block size in bits.
' @remarks This will always return 64 bits (8 bytes).
'
Public Property Get BlockSize() As Long
    BlockSize = mBase.BlockSize
End Property

''
' Sets the block size, in bits, used by this algorithm.
'
' @param Value The new block size in bits.
' @remarks The only valid block size is 64 bits.
'
Public Property Let BlockSize(ByVal Value As Long)
    mBase.BlockSize = Value
End Property

''
' Returns a list of legal block sizes.
'
' @return A list of legal block sizes.
' @remarks Only one block size is legal, 64 bits.
'
Public Property Get LegalBlockSizes() As KeySizes()
    LegalBlockSizes = mBase.LegalBlockSizes
End Property

''
' Returns the Feedback Size in bits.
'
' @return The Feedback size in bits.
' @remarks Always returns 64 bits.
'
Public Property Get FeedbackSize() As Long
    FeedbackSize = mBase.FeedbackSize
End Property

''
' Sets the Feedback size in bits.
'
' @param Value The new feedback size in bits.
' @remarks Only 64 bits is valid.
'
Public Property Let FeedbackSize(ByVal Value As Long)
    mBase.FeedbackSize = Value
End Property

''
' Returns the key size in bits.
'
' @Return The key size in bits.
'
Public Property Get KeySize() As Long
    KeySize = mBase.KeySize
End Property

''
' Sets the key size in bits.
'
' @param Value The new key size in bits.
' @remarks Key sizes of 128 or 192 are supported.
'
Public Property Let KeySize(ByVal Value As Long)
    mBase.KeySize = Value
End Property

''
' Returns a list of valid key sizes supported by this algorithm.
'
' @return A list of valid key sizes.
'
Public Property Get LegalKeySizes() As KeySizes()
    LegalKeySizes = mBase.LegalKeySizes
End Property

''
' Returns the Cipher mode this algorithm will use.
'
' @return The Cipher mode.
' @see CipherMode
'
Public Property Get Mode() As CipherMode
    Mode = mBase.Mode
End Property

''
' Sets the Cipher mode this algorithm will use.
'
' @param Value The cipher mode to use.
' @see CipherMode
'
Public Property Let Mode(ByVal Value As CipherMode)
    mBase.Mode = Value
End Property

''
' Returns the padding mode being used.
'
' @return The current padding mode.
' @see PaddingMode
'
Public Property Get Padding() As PaddingMode
    Padding = mBase.Padding
End Property

''
' Sets the padding mode to be used.
'
' @param Value The new padding mode.
' @see PaddingMode
'
Public Property Let Padding(ByVal Value As PaddingMode)
    mBase.Padding = Value
End Property

''
' Returns the current secret key.
'
' @return The secret key.
' @remarks If no key has been set, then a new random key will be generated.
'
Public Property Get Key() As Byte()
    Key = mBase.GetKey
End Property

''
' Sets the secret key.
'
' @param Value The new secret key.
' @param The key must be a valid key length. Currently only 64 bits (8 bytes) is supported.
'
Public Property Let Key(ByRef Value() As Byte)
    mBase.SetKey Value
End Property

''
' Generates a random key to be used.
'
Public Sub GenerateKey()
    mBase.GenerateKey
End Sub

''
' Returns the Initialization Vector.
'
' @return The Initialization Vector.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same as well. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using the <b>GenerateIV</b> method, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
' <p>If no IV has been set, then a new random IV will be generated.</p>
'
Public Property Get IV() As Byte()
    IV = mBase.GetIV
End Property

''
' Sets the Initialization Vector to be used.
'
' @param Value The new IV to use.
' @remarks If several blocks of same plain text are encrypted, the ciphered data could all be the same as well. Using an
' Initialization Vector can prevent same plain text blocks to reproduce the same ciphered blocks.
' <p>This can be generated using the <b>GenerateIV</b> method, or a new one will be generated if one hasn't been set when
' this property is accessed.</p>
' <p>The array must be 8 bytes in length.</p>
'
Public Property Let IV(ByRef Value() As Byte)
    mBase.SetIV Value
End Property

''
' Generates a random IV array to use.
'
Public Sub GenerateIV()
    mBase.GenerateIV
End Sub

''
' Clears the Key and IV arrays.
'
Public Sub Clear()
    mBase.Clear
End Sub

''
' Tests if a specific bit length is valid for a key.
'
' @param BitLength The number of bits to test for.
' @return Return True if the number of bits is supported, False otherwise.
' @remarks Only 128bit and 192bit are supported.
'
Public Function ValidKeySize(ByVal BitLength As Long) As Boolean
    ValidKeySize = mBase.ValidKeySize(BitLength)
End Function

''
' Returns a new cipher used to encrypt data.
'
' @param RgbKey A supplied byte array used as the secret key.
' @param RgbIV A supplied byte array used for the Initialization Vector (IV).
' @return An ICryptoTransform object used to encrypt data.
' @remarks Both parameters must be either supplied or missing. A single parameter will throw an exception.
' <p>If no parameters are supplied then the key and IV will be retrieved through the Key and IV properties.</p>
'
Public Function CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set CreateEncryptor = CreateCipher(RgbKey, RgbIV, True)
End Function

''
' Returns a new cipher used to decrypt data.
'
' @param RgbKey A supplied byte array used as the secret key.
' @param RgbIV A supplied byte array used for the Initialization Vector (IV).
' @return An ICryptoTransform object used to dencrypt data.
' @remarks Both parameters must be either supplied or missing. A single parameter will throw an exception.
' <p>If no parameters are supplied then the key and IV will be retrieved through the Key and IV properties.</p>
'
Public Function CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set CreateDecryptor = CreateCipher(RgbKey, RgbIV, False)
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
Private Function CreateCipher(ByRef RgbKey As Variant, ByRef RgbIV As Variant, ByVal IsEncrypting As Boolean) As CryptoAPITransform
    If Not IsMissing(RgbKey) Then
        If Mode <> CipherMode.ECB And IsMissing(RgbIV) Then
            Throw Cor.NewArgumentException("IV must be supplied with Key.", "RgbIV")
        End If
    ElseIf Not IsMissing(RgbIV) Then
        Throw Cor.NewArgumentException("Key must be supplied with IV.", "RgbKey")
    End If
    
    If Mode = CipherMode.CFB And FeedbackSize <> DefaultFeedbackSize Then _
        Error.Cryptographic Cryptography_CSP_CFBSizeNotSupported
    
    Dim Alg As Long
    If mBase.KeySize = 128 Then
        Alg = CALG_3DES_112
    Else
        Alg = CALG_3DES
    End If

    Set CreateCipher = New CryptoAPITransform
    CreateCipher.Init Alg, mBase.CloneRgbKey(RgbKey), mBase.CloneRgbIV(RgbIV), mBase, IsEncrypting
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = New SymmetricAlgorithmBase
    With mBase
        .AddLegalBlockSize LegalBlockSize, LegalBlockSize, 0
        .AddLegalKeySize LegalMinKeySize, LegalMaxKeySize, LegalSkipSize
        .BlockSize = LegalBlockSize
        .KeySize = DefaultKeySize
        .FeedbackSize = DefaultFeedbackSize
        .Mode = CipherMode.CBC
        .Padding = PaddingMode.PKCS7
    End With
End Sub


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
'   mBase Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mBase_CheckForWeakKey(RgbKey() As Byte, Ex As Exception)
    If Statics.TripleDES.IsWeakKey(RgbKey) Then
        Set Ex = Cor.NewCryptographicException("Key cannot be weak.")
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SymmetricAlgorithm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let SymmetricAlgorithm_BlockSize(ByVal Value As Long)
    BlockSize = Value
End Property

Private Property Get SymmetricAlgorithm_BlockSize() As Long
    SymmetricAlgorithm_BlockSize = BlockSize
End Property

Private Sub SymmetricAlgorithm_Clear()
    Clear
End Sub

Private Function SymmetricAlgorithm_CreateDecryptor(Optional RgbKey As Variant, Optional RgbIV As Variant) As ICryptoTransform
    Set SymmetricAlgorithm_CreateDecryptor = CreateDecryptor(RgbKey, RgbIV)
End Function

Private Function SymmetricAlgorithm_CreateEncryptor(Optional RgbKey As Variant, Optional RgbIV As Variant) As ICryptoTransform
    Set SymmetricAlgorithm_CreateEncryptor = CreateEncryptor(RgbKey, RgbIV)
End Function

Private Function SymmetricAlgorithm_Equals(Value As Variant) As Boolean
    SymmetricAlgorithm_Equals = Equals(Value)
End Function

Private Property Let SymmetricAlgorithm_FeedbackSize(ByVal Value As Long)
    FeedbackSize = Value
End Property

Private Property Get SymmetricAlgorithm_FeedbackSize() As Long
    SymmetricAlgorithm_FeedbackSize = FeedbackSize
End Property

Private Sub SymmetricAlgorithm_GenerateIV()
    GenerateIV
End Sub

Private Sub SymmetricAlgorithm_GenerateKey()
    GenerateKey
End Sub

Private Function SymmetricAlgorithm_GetHashCode() As Long
    SymmetricAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Let SymmetricAlgorithm_IV(Value() As Byte)
    mBase.SetIV Value
End Property

Private Property Get SymmetricAlgorithm_IV() As Byte()
    SymmetricAlgorithm_IV = IV
End Property

Private Property Let SymmetricAlgorithm_Key(Value() As Byte)
    mBase.SetKey Value
End Property

Private Property Get SymmetricAlgorithm_Key() As Byte()
    SymmetricAlgorithm_Key = Key
End Property

Private Property Let SymmetricAlgorithm_KeySize(ByVal Value As Long)
    KeySize = Value
End Property

Private Property Get SymmetricAlgorithm_KeySize() As Long
    SymmetricAlgorithm_KeySize = KeySize
End Property

Private Property Get SymmetricAlgorithm_LegalBlockSizes() As KeySizes()
    SymmetricAlgorithm_LegalBlockSizes = LegalBlockSizes
End Property

Private Property Get SymmetricAlgorithm_LegalKeySizes() As KeySizes()
    SymmetricAlgorithm_LegalKeySizes = LegalKeySizes
End Property

Private Property Let SymmetricAlgorithm_Mode(ByVal Value As CipherMode)
    Mode = Value
End Property

Private Property Get SymmetricAlgorithm_Mode() As CipherMode
    SymmetricAlgorithm_Mode = Mode
End Property

Private Property Let SymmetricAlgorithm_Padding(ByVal Value As PaddingMode)
    Padding = Value
End Property

Private Property Get SymmetricAlgorithm_Padding() As PaddingMode
    SymmetricAlgorithm_Padding = Padding
End Property

Private Function SymmetricAlgorithm_ToString() As String
    SymmetricAlgorithm_ToString = ToString
End Function

Private Function SymmetricAlgorithm_ValidKeySize(ByVal BitLength As Long) As Boolean
    SymmetricAlgorithm_ValidKeySize = ValidKeySize(BitLength)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   TripleDES
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let TripleDES_BlockSize(ByVal Value As Long)
    BlockSize = Value
End Property

Private Property Get TripleDES_BlockSize() As Long
    TripleDES_BlockSize = BlockSize
End Property

Private Sub TripleDES_Clear()
    Clear
End Sub

Private Function TripleDES_CreateDecryptor(Optional RgbKey As Variant, Optional RgbIV As Variant) As ICryptoTransform
    Set TripleDES_CreateDecryptor = CreateDecryptor(RgbKey, RgbIV)
End Function

Private Function TripleDES_CreateEncryptor(Optional RgbKey As Variant, Optional RgbIV As Variant) As ICryptoTransform
    Set TripleDES_CreateEncryptor = CreateEncryptor(RgbKey, RgbIV)
End Function

Private Function TripleDES_Equals(Value As Variant) As Boolean
    TripleDES_Equals = Equals(Value)
End Function

Private Property Let TripleDES_FeedbackSize(ByVal Value As Long)
    FeedbackSize = Value
End Property

Private Property Get TripleDES_FeedbackSize() As Long
    TripleDES_FeedbackSize = FeedbackSize
End Property

Private Sub TripleDES_GenerateIV()
    GenerateIV
End Sub

Private Sub TripleDES_GenerateKey()
    GenerateKey
End Sub

Private Function TripleDES_GetHashCode() As Long
    TripleDES_GetHashCode = GetHashCode
End Function

Private Property Let TripleDES_IV(Value() As Byte)
    mBase.SetIV Value
End Property

Private Property Get TripleDES_IV() As Byte()
    TripleDES_IV = IV
End Property

Private Property Let TripleDES_Key(Value() As Byte)
    mBase.SetKey Value
End Property

Private Property Get TripleDES_Key() As Byte()
    TripleDES_Key = Key
End Property

Private Property Let TripleDES_KeySize(ByVal Value As Long)
    KeySize = Value
End Property

Private Property Get TripleDES_KeySize() As Long
    TripleDES_KeySize = KeySize
End Property

Private Property Get TripleDES_LegalBlockSizes() As KeySizes()
    TripleDES_LegalBlockSizes = LegalBlockSizes
End Property

Private Property Get TripleDES_LegalKeySizes() As KeySizes()
    TripleDES_LegalKeySizes = LegalKeySizes
End Property

Private Property Let TripleDES_Mode(ByVal Value As CipherMode)
    Mode = Value
End Property

Private Property Get TripleDES_Mode() As CipherMode
    TripleDES_Mode = Mode
End Property

Private Property Let TripleDES_Padding(ByVal Value As PaddingMode)
    Padding = Value
End Property

Private Property Get TripleDES_Padding() As PaddingMode
    TripleDES_Padding = Padding
End Property

Private Function TripleDES_ToString() As String
    TripleDES_ToString = ToString
End Function

Private Function TripleDES_ValidKeySize(ByVal BitLength As Long) As Boolean
    TripleDES_ValidKeySize = ValidKeySize(BitLength)
End Function

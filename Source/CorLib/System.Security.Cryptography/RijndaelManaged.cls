VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RijndaelManaged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: RijndaelManaged
'

''
' Represents an implementation of the Rijndael symmetric block cipher algorithm
'
' @remarks Details about Rijndael can be found here: http://www.iaik.tugraz.at/research/krypto/AES/old/~rijmen/rijndael/
' @see RijndaelManagedTransform
' @include "..\DocIncludes\RijndaelManaged.html"
Option Explicit
Implements IObject
Implements SymmetricAlgorithm
Implements Rijndael

Private Const DEF_MINBLOCKSIZE  As Long = 128
Private Const DEF_MAXBLOCKSIZE  As Long = 256
Private Const DEF_BLOCKSIZE     As Long = DEF_MINBLOCKSIZE
Private Const DEF_BLOCKSKIPSIZE As Long = 64
Private Const DEF_FEEDBACKSIZE  As Long = DEF_BLOCKSIZE
Private Const DEF_MINKEYSIZE    As Long = 128
Private Const DEF_MAXKEYSIZE    As Long = 256
Private Const DEF_KEYSIZE       As Long = DEF_MAXKEYSIZE
Private Const DEF_KEYSKIPSIZE   As Long = 64

Private mBase As SymmetricAlgorithmBase
Attribute mBase.VB_VarHelpID = -1


''
' Returns the block size, in bits, used by this algorithm.
'
' @return The block size in bits.
' @remarks Block sizes of 128, 192, and 256 are supported.
'
Public Property Get BlockSize() As Long
    BlockSize = mBase.BlockSize
End Property

''
' Sets the block size, in bits, used by this algorithm.
'
' @param Value The new block size in bits.
' @remarks Block sizes of 128, 192, and 256 are supported.
'
Public Property Let BlockSize(ByVal Value As Long)
    mBase.BlockSize = Value
End Property

''
' Returns a list of legal block sizes.
'
' @return A list of legal block sizes.
' @remarks Block sizes of 128, 192, and 256 are supported.
'
Public Property Get LegalBlockSizes() As KeySizes()
    LegalBlockSizes = mBase.LegalBlockSizes
End Property

''
' Returns the Feedback Size in bits.
'
' @return The Feedback size in bits.
'
Public Property Get FeedbackSize() As Long
    FeedbackSize = mBase.FeedbackSize
End Property

''
' Sets the Feedback size in bits.
'
' @param Value The new feedback size in bits.
'
Public Property Let FeedbackSize(ByVal Value As Long)
    mBase.FeedbackSize = Value
End Property

''
' Returns the key size in bits.
'
' @Return The key size in bits.
' @remarks 128, 192, and 256 bits are supported.
'
Public Property Get KeySize() As Long
    KeySize = mBase.KeySize
End Property

''
' Sets the key size in bits.
'
' @param Value The new key size in bits.
' @remarks 128, 192, and 256 bits are supported.
'
Public Property Let KeySize(ByVal Value As Long)
    mBase.KeySize = Value
End Property

''
' Returns a list of valid key sizes supported by this algorithm.
'
' @return A list of valid key sizes.
' @remarks Key sizes of 128, 192, and 256 are supported.
'
Public Property Get LegalKeySizes() As KeySizes()
    LegalKeySizes = mBase.LegalKeySizes
End Property

''
' Returns the Cipher mode this algorithm will use.
'
' @return The Cipher mode.
'
Public Property Get Mode() As CipherMode
    Mode = mBase.Mode
End Property

''
' Sets the Cipher mode this algorithm will use.
'
' @param Value The cipher mode to use.
'
Public Property Let Mode(ByVal Value As CipherMode)
    mBase.Mode = Value
End Property

''
' Returns the padding mode being used.
'
' @return The current padding mode.
'
Public Property Get Padding() As PaddingMode
    Padding = mBase.Padding
End Property

''
' Sets the padding mode to be used.
'
' @param Value The new padding mode.
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
' @param The key must be a valid key length.
' @remarks  Key sizes of 128, 192, and 256 bits are supported.
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
' @remarks If no IV has been set, then a new random IV will be generated.
'
Public Property Get IV() As Byte()
    IV = mBase.GetIV
End Property

''
' Sets the Initialization Vector to be used.
'
' @param Value The new IV to use.
' @remarks The array must be 8 bytes in length.
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
' @remarks 128, 192, 256 bit keys are supported.
'
Public Function ValidKeySize(ByVal BitLength As Long) As Boolean
    ValidKeySize = mBase.ValidKeySize(BitLength)
End Function

''
' Returns a new cipher used to encrypt data.
'
' @param RgbKey A supplied byte array used as the secret key.
' @param RgbIV A supplied byte array used for the Initialization Vector (IV).
' @return An ICryptoTransform (RijndaelManagedTransform) object used to encrypt data.
' @remarks Both parameters must be either supplied or missing. A single parameter will throw an exception.
' <p>If no parameters are supplied then the key and IV will be retrieved through the Key and IV properties.</p>
' @see RijndaelManagedTransform
'
Public Function CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set CreateEncryptor = CreateCipher(RgbKey, RgbIV, True)
End Function

''
' Returns a new cipher used to decrypt data.
'
' @param RgbKey A supplied byte array used as the secret key.
' @param RgbIV A supplied byte array used for the Initialization Vector (IV).
' @return An ICryptoTransform (RijndaelManagedTransform) object used to dencrypt data.
' @remarks Both parameters must be either supplied or missing. A single parameter will throw an exception.
' <p>If no parameters are supplied then the key and IV will be retrieved through the Key and IV properties.</p>
' @see RijndaelManagedTransform
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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateCipher(ByRef RgbKey As Variant, ByRef RgbIV As Variant, ByVal IsEncrypting As Boolean) As RijndaelManagedTransform
    If IsMissing(RgbKey) <> IsMissing(RgbIV) Then _
        Throw Cor.NewArgumentException("Argument is missing.", IIf(IsMissing(RgbKey), "RgbKey", "RgbIV"))
    
    Set CreateCipher = New RijndaelManagedTransform
    CreateCipher.Init mBase.CloneRgbKey(RgbKey), mBase.CloneRgbIV(RgbIV), mBase, IsEncrypting
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = New SymmetricAlgorithmBase
    With mBase
        .AddLegalBlockSize DEF_MINBLOCKSIZE, DEF_MAXBLOCKSIZE, DEF_BLOCKSKIPSIZE
        .AddLegalKeySize DEF_MINKEYSIZE, DEF_MAXKEYSIZE, DEF_KEYSKIPSIZE
        .BlockSize = DEF_BLOCKSIZE
        .KeySize = DEF_KEYSIZE
        .FeedbackSize = DEF_FEEDBACKSIZE
        .Mode = CipherMode.CBC
        .Padding = PaddingMode.PKCS7
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Rijndael
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let Rijndael_BlockSize(ByVal Value As Long)
    BlockSize = Value
End Property

Private Property Get Rijndael_BlockSize() As Long
    Rijndael_BlockSize = BlockSize
End Property

Private Sub Rijndael_Clear()
    Clear
End Sub

Private Function Rijndael_CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set Rijndael_CreateDecryptor = CreateDecryptor(RgbKey, RgbIV)
End Function

Private Function Rijndael_CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set Rijndael_CreateEncryptor = CreateEncryptor(RgbKey, RgbIV)
End Function

Private Function Rijndael_Equals(Value As Variant) As Boolean
    Rijndael_Equals = Equals(Value)
End Function

Private Property Let Rijndael_FeedbackSize(ByVal Value As Long)
    FeedbackSize = Value
End Property

Private Property Get Rijndael_FeedbackSize() As Long
    Rijndael_FeedbackSize = FeedbackSize
End Property

Private Sub Rijndael_GenerateIV()
    GenerateIV
End Sub

Private Sub Rijndael_GenerateKey()
    GenerateKey
End Sub

Private Function Rijndael_GetHashCode() As Long
    Rijndael_GetHashCode = GetHashCode
End Function

Private Property Let Rijndael_IV(Value() As Byte)
    mBase.SetIV Value
End Property

Private Property Get Rijndael_IV() As Byte()
    Rijndael_IV = IV
End Property

Private Property Let Rijndael_Key(Value() As Byte)
    mBase.SetKey Value
End Property

Private Property Get Rijndael_Key() As Byte()
    Rijndael_Key = Key
End Property

Private Property Let Rijndael_KeySize(ByVal Value As Long)
    KeySize = Value
End Property

Private Property Get Rijndael_KeySize() As Long
    Rijndael_KeySize = KeySize
End Property

Private Property Get Rijndael_LegalBlockSizes() As KeySizes()
    Rijndael_LegalBlockSizes = LegalBlockSizes
End Property

Private Property Get Rijndael_LegalKeySizes() As KeySizes()
    Rijndael_LegalKeySizes = LegalKeySizes
End Property

Private Property Let Rijndael_Mode(ByVal Value As CipherMode)
    Mode = Value
End Property

Private Property Get Rijndael_Mode() As CipherMode
    Rijndael_Mode = Mode
End Property

Private Property Let Rijndael_Padding(ByVal Value As PaddingMode)
    Padding = Value
End Property

Private Property Get Rijndael_Padding() As PaddingMode
    Rijndael_Padding = Padding
End Property

Private Function Rijndael_ToString() As String
    Rijndael_ToString = ToString
End Function

Private Function Rijndael_ValidKeySize(ByVal BitLength As Long) As Boolean
    Rijndael_ValidKeySize = ValidKeySize(BitLength)
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

Private Function SymmetricAlgorithm_CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set SymmetricAlgorithm_CreateDecryptor = CreateDecryptor(RgbKey, RgbIV)
End Function

Private Function SymmetricAlgorithm_CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
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


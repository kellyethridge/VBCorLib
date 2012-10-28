VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RC2CryptoServiceProvider"
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
'    Module: RC2CryptoServiceProvider
'

''
' Provides cryptography services using the RC2 symmetric algorithm.
'
' @remarks This class uses the Windows CryptoAPI services.
' @see RC2
' @include "..\..\..\Includes\RC2CryptoServiceProviderExample.txt"
Option Explicit
Implements IObject
Implements SymmetricAlgorithm
Implements RC2

Private Const DEF_BLOCKSIZE     As Long = 64
Private Const DEF_MINBLOCKSIZE  As Long = DEF_BLOCKSIZE
Private Const DEF_MAXBLOCKSIZE  As Long = DEF_BLOCKSIZE
Private Const DEF_BLOCKSKIPSIZE As Long = 0
Private Const DEF_FEEDBACKSIZE  As Long = 8
Private Const DEF_MINKEYSIZE    As Long = 40
Private Const DEF_MAXKEYSIZE    As Long = 128
Private Const DEF_KEYSIZE       As Long = DEF_MAXKEYSIZE
Private Const DEF_KEYSKIPSIZE   As Long = 8

Private mBase       As SymmetricAlgorithmBase
Attribute mBase.VB_VarHelpID = -1
Private mUseSalt    As Boolean



''
' Returns the effective key size to be used.
'
' @return The effective key size.
'
Public Property Get EffectiveKeySize() As Long
    EffectiveKeySize = mBase.KeySize
End Property

''
' Sets the effective key size to be used.
'
' @param RHS The new effective key size.
' @remarks This property can only be set to the same value as KeySize.
' This property is useless in setting the effective key size. It is here
' to complete the .NET interface for this class.
' <p>The effective key size to use will ultimately be determined by
' the length of the actual key being used.</p>
'
Public Property Let EffectiveKeySize(ByVal RHS As Long)
    If RHS <> mBase.KeySize Then Throw Cor.NewCryptographicException("EffectiveKeySize must match KeySize.")
End Property

''
' Returns if a salt value is to be used with the key generation.
'
' @return Returns if the a salt value will be used.
'
Public Property Get UseSalt() As Boolean
    UseSalt = mUseSalt
End Property

''
' Sets if a salt value will be used with the key generation.
'
' @param RHS Whether a salt value will be used with the key generation.
'
Public Property Let UseSalt(ByVal RHS As Boolean)
    mUseSalt = RHS
End Property

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
' @param RHS The new block size in bits.
' @remarks The only valid block size is 64 bits.
'
Public Property Let BlockSize(ByVal RHS As Long)
    mBase.BlockSize = RHS
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
' @param RHS The new feedback size in bits.
' @remarks Only 64 bits is valid.
'
Public Property Let FeedbackSize(ByVal RHS As Long)
    mBase.FeedbackSize = RHS
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
' @param RHS The new key size in bits.
' @remarks Key sizes of 128 or 192 are supported.
'
Public Property Let KeySize(ByVal RHS As Long)
    mBase.KeySize = RHS
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
'
Public Property Get Mode() As CipherMode
    Mode = mBase.Mode
End Property

''
' Sets the Cipher mode this algorithm will use.
'
' @param RHS The cipher mode to use.
'
Public Property Let Mode(ByVal RHS As CipherMode)
    mBase.Mode = RHS
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
' @param RHS The new padding mode.
'
Public Property Let Padding(ByVal RHS As PaddingMode)
    mBase.Padding = RHS
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
' @param RHS The new secret key.
' @param The key must be a valid key length. Currently only 64 bits (8 bytes) is supported.
'
Public Property Let Key(ByRef RHS() As Byte)
    Call mBase.SetKey(RHS)
End Property

''
' Generates a random key to be used.
'
Public Sub GenerateKey()
    Call mBase.GenerateKey
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
' @param RHS The new IV to use.
' @remarks The array must be 8 bytes in length.
'
Public Property Let IV(ByRef RHS() As Byte)
    Call mBase.SetIV(RHS)
End Property

''
' Generates a random IV array to use.
'
Public Sub GenerateIV()
    Call mBase.GenerateIV
End Sub

''
' Clears the Key and IV arrays.
'
Public Sub Clear()
    Call mBase.Clear
End Sub

''
' Tests if a specific bit length is valid for a key.
'
' @param BitLength The number of bits to test for.
' @return Return True if the number of bits is supported, False otherwise.
' @remarks Valid sizes are from 40 bits to 128 bits in steps of 8 bits.
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
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CreateCipher(ByRef RgbKey As Variant, ByRef RgbIV As Variant, ByVal IsEncrypting As Boolean) As CryptoAPITransform
    If IsMissing(RgbKey) <> IsMissing(RgbIV) Then _
        Throw Cor.NewArgumentException("Argument is missing.", IIf(IsMissing(RgbKey), "RgbKey", "RgbIV"))
    
    Dim Key() As Byte
    Key = mBase.CloneRgbKey(RgbKey)
    
    ' We calculate the effective key size because the key itself may have been passed into the CreateEncryptor/CreateDecryptor
    ' functions so the KeySize property may not reflect the true size of the key to be used.
    Set CreateCipher = New CryptoAPITransform
    Call CreateCipher.Init(CALG_RC2, Key, mBase.CloneRgbIV(RgbIV), mBase, IsEncrypting, cArray.GetLength(Key) * 8, mUseSalt)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set mBase = New SymmetricAlgorithmBase
    With mBase
        Call .AddLegalBlockSize(DEF_MINBLOCKSIZE, DEF_MAXBLOCKSIZE, DEF_BLOCKSKIPSIZE)
        Call .AddLegalKeySize(DEF_MINKEYSIZE, DEF_MAXKEYSIZE, DEF_KEYSKIPSIZE)
        .BlockSize = DEF_BLOCKSIZE
        .KeySize = DEF_KEYSIZE
        .FeedbackSize = DEF_FEEDBACKSIZE
        .Mode = CipherMode.CBC
        .Padding = PaddingMode.PKCS7
    End With
End Sub


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
'   RC2 Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let RC2_BlockSize(ByVal RHS As Long)
    BlockSize = RHS
End Property

Private Property Get RC2_BlockSize() As Long
    RC2_BlockSize = BlockSize
End Property

Private Sub RC2_Clear()
    Call Clear
End Sub

Private Function RC2_CreateDecryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set RC2_CreateDecryptor = CreateDecryptor(RgbKey, RgbIV)
End Function

Private Function RC2_CreateEncryptor(Optional ByRef RgbKey As Variant, Optional ByRef RgbIV As Variant) As ICryptoTransform
    Set RC2_CreateEncryptor = CreateEncryptor(RgbKey, RgbIV)
End Function

Private Property Let RC2_EffectiveKeySize(ByVal RHS As Long)
    EffectiveKeySize = RHS
End Property

Private Property Get RC2_EffectiveKeySize() As Long
    RC2_EffectiveKeySize = EffectiveKeySize
End Property

Private Function RC2_Equals(Value As Variant) As Boolean
    RC2_Equals = Equals(Value)
End Function

Private Property Let RC2_FeedbackSize(ByVal RHS As Long)
    FeedbackSize = RHS
End Property

Private Property Get RC2_FeedbackSize() As Long
    RC2_FeedbackSize = FeedbackSize
End Property

Private Sub RC2_GenerateIV()
    Call GenerateIV
End Sub

Private Sub RC2_GenerateKey()
    Call GenerateKey
End Sub

Private Function RC2_GetHashCode() As Long
    RC2_GetHashCode = GetHashCode
End Function

Private Property Let RC2_IV(RHS() As Byte)
    Call mBase.SetIV(RHS)
End Property

Private Property Get RC2_IV() As Byte()
    RC2_IV = IV
End Property

Private Property Let RC2_Key(RHS() As Byte)
    Call mBase.SetKey(RHS)
End Property

Private Property Get RC2_Key() As Byte()
    RC2_Key = Key
End Property

Private Property Let RC2_KeySize(ByVal RHS As Long)
    KeySize = RHS
End Property

Private Property Get RC2_KeySize() As Long
    RC2_KeySize = KeySize
End Property

Private Property Get RC2_LegalBlockSizes() As KeySizes()
    RC2_LegalBlockSizes = LegalBlockSizes
End Property

Private Property Get RC2_LegalKeySizes() As KeySizes()
    RC2_LegalKeySizes = LegalKeySizes
End Property

Private Property Let RC2_Mode(ByVal RHS As CipherMode)
    Mode = RHS
End Property

Private Property Get RC2_Mode() As CipherMode
    RC2_Mode = Mode
End Property

Private Property Let RC2_Padding(ByVal RHS As PaddingMode)
    Padding = RHS
End Property

Private Property Get RC2_Padding() As PaddingMode
    RC2_Padding = Padding
End Property

Private Function RC2_ToString() As String
    RC2_ToString = ToString
End Function

Private Function RC2_ValidKeySize(ByVal BitLength As Long) As Boolean
    RC2_ValidKeySize = ValidKeySize(BitLength)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   SymmetricAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let SymmetricAlgorithm_BlockSize(ByVal RHS As Long)
    BlockSize = RHS
End Property

Private Property Get SymmetricAlgorithm_BlockSize() As Long
    SymmetricAlgorithm_BlockSize = BlockSize
End Property

Private Sub SymmetricAlgorithm_Clear()
    Call Clear
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

Private Property Let SymmetricAlgorithm_FeedbackSize(ByVal RHS As Long)
    FeedbackSize = RHS
End Property

Private Property Get SymmetricAlgorithm_FeedbackSize() As Long
    SymmetricAlgorithm_FeedbackSize = FeedbackSize
End Property

Private Sub SymmetricAlgorithm_GenerateIV()
    Call GenerateIV
End Sub

Private Sub SymmetricAlgorithm_GenerateKey()
    Call GenerateKey
End Sub

Private Function SymmetricAlgorithm_GetHashCode() As Long
    SymmetricAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Let SymmetricAlgorithm_IV(RHS() As Byte)
    Call mBase.SetIV(RHS)
End Property

Private Property Get SymmetricAlgorithm_IV() As Byte()
    SymmetricAlgorithm_IV = IV
End Property

Private Property Let SymmetricAlgorithm_Key(RHS() As Byte)
    Call mBase.SetKey(RHS)
End Property

Private Property Get SymmetricAlgorithm_Key() As Byte()
    SymmetricAlgorithm_Key = Key
End Property

Private Property Let SymmetricAlgorithm_KeySize(ByVal RHS As Long)
    KeySize = RHS
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

Private Property Let SymmetricAlgorithm_Mode(ByVal RHS As CipherMode)
    Mode = RHS
End Property

Private Property Get SymmetricAlgorithm_Mode() As CipherMode
    SymmetricAlgorithm_Mode = Mode
End Property

Private Property Let SymmetricAlgorithm_Padding(ByVal RHS As PaddingMode)
    Padding = RHS
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


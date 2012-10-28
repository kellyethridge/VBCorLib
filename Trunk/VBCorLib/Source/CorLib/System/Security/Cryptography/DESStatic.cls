VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DESStatic"
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
'    Module: DESStatic
'

''
' Provides static methods relating to the DES crypto service provider.
'
' @see DES
' @see DESCryptoServiceProvider
'
Option Explicit

Private Const DES_KEY_LENGTH    As Long = 8    ' 56 bits + 8 bits parity
Private Const DEF_BLOCKSIZE     As Long = 64
Private Const DEF_KEYSIZE       As Long = 64

Private mWeakKeys()         As Currency
Private mSemiWeakKeys()     As Currency
Private mLegalKeySizes()    As KeySizes
Private mLegalBlockSizes()  As KeySizes



''
' Determins if a key is considered semi-weak.
'
' @param RgbKey The key to check for semi-weakness.
' @return Returns True if the key is semi-weak, False otherwise.
' @remarks DES has a set of keys known to be semi-weak. The key to be tested
' is compared to this set of known semi-weak keys.
' <p>The set of semi-weak keys are:
' <pre>
'    "E001E001F101F101"
'    "01E001E001F101F1"
'    "FE1FFE1FFE0EFE0E"
'    "1FFE1FFE0EFE0EFE"
'    "E01FE01FF10EF10E"
'    "1FE01FE00EF10EF1"
'    "01FE01FE01FE01FE"
'    "FE01FE01FE01FE01"
'    "011F011F010E010E"
'    "1F011F010E010E01"
'    "E0FEE0FEF1FEF1FE"
'    "FEE0FEE0FEF1FEF1"
' </pre>
'
Public Function IsSemiWeakKey(ByRef RgbKey() As Byte) As Boolean
    IsSemiWeakKey = MatchesKnownKey(RgbKey, mSemiWeakKeys)
End Function

''
' Determines if a key is considered weak.
'
' @param RgbKey The key to check for weakness.
' @return Returns True if the key weak, False otherwise.
' @remarks DES has a set of keys known to be weak. The key to be tested
' is compared against this set of known weak keys.
' <p>The set of weak keys are:
' <pre>
'    "0101010101010101"
'    "FEFEFEFEFEFEFEFE"
'    "E0E0E0E0F1F1F1F1"
'    "1F1F1F1F0E0E0E0E"
' </pre></p>
' <p>What makes these keys weak is if test encrypted with one of these keys
' is encrypted a second time, the original plain text would be produced.</p>
'
Public Function IsWeakKey(ByRef RgbKey() As Byte) As Boolean
    IsWeakKey = MatchesKnownKey(RgbKey, mWeakKeys)
End Function

''
' Returns a DES cryptography service provider based on the name of the provider.
'
' @param AlgName The implementation name of the DES cryptography service provider.
' @return Returns a DES implementation.
' @remarks If no name is given, then the default implementation is returned.
'
Public Function Create(Optional ByVal AlgName As String = "DES") As DES
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyKeySize(ByRef RgbKey As Variant)
    If cArray.IsNull(RgbKey) Then _
        Throw Cor.NewCryptographicException("Invalid key length.")
    
    If cArray.GetLength(RgbKey) <> DES_KEY_LENGTH Then _
        Throw Cor.NewCryptographicException("Invalid key length.")
End Sub

Private Function MatchesKnownKey(ByRef RgbKey() As Byte, ByRef WeakKeys() As Currency) As Boolean
    Dim Key As Currency
    Key = Get64BitKey(RgbKey)
    
    Dim i As Long
    For i = 0 To UBound(WeakKeys)
        If WeakKeys(i) = Key Then
            MatchesKnownKey = True
            Exit Function
        End If
    Next i
End Function

Private Sub InitWeakKeys()
    ReDim mWeakKeys(3)
    mWeakKeys(0) = Parse64BitKeyString("0101010101010101")
    mWeakKeys(1) = Parse64BitKeyString("FEFEFEFEFEFEFEFE")
    mWeakKeys(2) = Parse64BitKeyString("E0E0E0E0F1F1F1F1")
    mWeakKeys(3) = Parse64BitKeyString("1F1F1F1F0E0E0E0E")
End Sub

Private Sub InitSemiWeakKeys()
    ReDim mSemiWeakKeys(11)
    mSemiWeakKeys(0) = Parse64BitKeyString("E001E001F101F101")
    mSemiWeakKeys(1) = Parse64BitKeyString("01E001E001F101F1")
    mSemiWeakKeys(2) = Parse64BitKeyString("FE1FFE1FFE0EFE0E")
    mSemiWeakKeys(3) = Parse64BitKeyString("1FFE1FFE0EFE0EFE")
    mSemiWeakKeys(4) = Parse64BitKeyString("E01FE01FF10EF10E")
    mSemiWeakKeys(5) = Parse64BitKeyString("1FE01FE00EF10EF1")
    mSemiWeakKeys(6) = Parse64BitKeyString("01FE01FE01FE01FE")
    mSemiWeakKeys(7) = Parse64BitKeyString("FE01FE01FE01FE01")
    mSemiWeakKeys(8) = Parse64BitKeyString("011F011F010E010E")
    mSemiWeakKeys(9) = Parse64BitKeyString("1F011F010E010E01")
    mSemiWeakKeys(10) = Parse64BitKeyString("E0FEE0FEF1FEF1FE")
    mSemiWeakKeys(11) = Parse64BitKeyString("FEE0FEE0FEF1FEF1")
End Sub

Private Function Parse64BitKeyString(ByVal s As String) As Currency
    Dim b(7) As Byte
    Dim i As Long
    For i = 0 To 7
        b(i) = CByte("&h" & Mid$(s, 1 + i * 2, 2)) Or 1
    Next i
    
    Parse64BitKeyString = AsCurr(b(0))
End Function

Private Function Get64BitKey(ByRef RgbKey() As Byte) As Currency
    Call VerifyKeySize(RgbKey)
    
    Dim lb      As Long
    lb = LBound(RgbKey)
    
    Dim b(7)    As Byte
    Dim i       As Long
    For i = 0 To 7
        b(i) = RgbKey(i + lb) Or 1 ' we set the parity bit to on
    Next i
    
    Get64BitKey = AsCurr(b(0))
End Function

Private Sub InitLegalSizes()
    ReDim mLegalKeySizes(0)
    Set mLegalKeySizes(0) = Cor.NewKeySizes(DEF_KEYSIZE, DEF_KEYSIZE, 0)
    
    ReDim mLegalBlockSizes(0)
    Set mLegalBlockSizes(0) = Cor.NewKeySizes(DEF_BLOCKSIZE, DEF_BLOCKSIZE, 0)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWeakKeys
    Call InitSemiWeakKeys
    Call InitLegalSizes
End Sub

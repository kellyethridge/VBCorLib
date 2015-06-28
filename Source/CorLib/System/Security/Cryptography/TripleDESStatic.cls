VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TripleDESStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: TripleDESStatic
'

''
' Provides static methods used with the Triple Data Encryption Standard algorithm.
'
' @see TripleDES
' @see TripleDESCryptoServiceProvider
'
Option Explicit

Private mLegalBlockSizes()  As KeySizes
Private mLegalKeySizes()    As KeySizes



''
' Returns a TripleDES algorithm implementation based on the name.
'
' @param AlgName The name of the TripleDES implementation.
' @return The TripleDES algorithm.
' @remarks The name is specified by the Server.Class name. The CreateObject
' method is used to instantiate the object.
' <p>If the name is "3DES", "TripleDES" or missing then the default implementation is returned.
'
Public Function Create(Optional ByVal AlgName As String = "TripleDES") As TripleDES
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function

''
' Determines if a key is weak or not.
'
' @param RgbKey The key to be tested for weakness.
' @return Returns True if the key is weak, or False if not.
' @remarks A weak key is determined if the first 64 bits are equal to the second 64 bits
' for a 128 bit key. For a 192 bit key, weakness is determined if the first 64 bits is
' equal to the second 64 bits, or if the second 64 bits is equal to the third 64 bits.
'
Public Function IsWeakKey(ByRef RgbKey() As Byte) As Boolean
    Dim Size As Long
    Size = CorArray.Length(RgbKey)
    
    If Not IsLegalKeySize(Size * 8) Then _
        Throw Cor.NewCryptographicException("Invalid key length.")
    
    If Size = 16 Then
        IsWeakKey = IsWeakKey128(RgbKey)
    Else
        IsWeakKey = IsWeakKey192(RgbKey)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsLegalKeySize(ByVal Bits As Long) As Boolean
    Dim i As Long
    For i = 0 To UBound(mLegalKeySizes)
        If mLegalKeySizes(i).IsLegalSize(Bits) Then
            IsLegalKeySize = True
            Exit Function
        End If
    Next i
End Function

Private Function IsWeakKey128(ByRef RgbKey() As Byte) As Boolean
    IsWeakKey128 = IsMatchingParts(RgbKey, LBound(RgbKey), LBound(RgbKey) + 8)
End Function

Private Function IsWeakKey192(ByRef RgbKey() As Byte) As Boolean
    If IsWeakKey128(RgbKey) Then
        IsWeakKey192 = True
    ElseIf IsMatchingParts(RgbKey, LBound(RgbKey) + 8, LBound(RgbKey) + 16) Then
        IsWeakKey192 = True
    End If
End Function

Private Function IsMatchingParts(ByRef RgbKey() As Byte, ByVal Part1Index As Long, ByVal Part2Index As Long) As Boolean
    Dim i As Long
    For i = 0 To 7
        If (RgbKey(Part1Index + i) And &HFE) <> (RgbKey(Part2Index + i) And &HFE) Then Exit Function
    Next i
    
    IsMatchingParts = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ReDim mLegalBlockSizes(0)
    Set mLegalBlockSizes(0) = Cor.NewKeySizes(64, 64, 0)
    
    ReDim mLegalKeySizes(0)
    Set mLegalKeySizes(0) = Cor.NewKeySizes(128, 192, 64)
End Sub

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
'    Module: TripleDESStatic
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
    Size = cArray.GetLength(RgbKey)
    
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

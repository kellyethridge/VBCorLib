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

Private Const PartSize As Long = 8


''
' Returns a TripleDES algorithm implementation based on the name.
'
' @param AlgName The name of the <a href="TripleDES.html">TripleDES</a> implementation.
' @return The specified <a href="TripleDES.html">TripleDES</a> algorithm.
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
    Const Key128Bit As Long = 16
    Const Key192Bit As Long = 24
    
    Dim Size As Long
    Size = CorArray.Length(RgbKey)
    
    Select Case Size
        Case Key128Bit
            IsWeakKey = HasMatchingParts(RgbKey, LBound(RgbKey))
        Case Key192Bit
            IsWeakKey = HasMatchingParts(RgbKey, LBound(RgbKey)) Or _
                        HasMatchingParts(RgbKey, LBound(RgbKey) + PartSize)
        Case Else
            Throw Cor.NewCryptographicException("Invalid key length.")
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasMatchingParts(ByRef RgbKey() As Byte, ByVal StartIndex As Long) As Boolean
    Dim i As Long
    
    For i = 0 To PartSize - 1
        If (RgbKey(StartIndex + i) And &HFE) <> (RgbKey(StartIndex + i + PartSize) And &HFE) Then
            Exit Function
        End If
    Next
    
    HasMatchingParts = True
End Function


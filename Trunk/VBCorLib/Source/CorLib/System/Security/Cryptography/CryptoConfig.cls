VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoConfig"
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
'    Module: CryptoConfig
'

''
' Used to access cryptography information.

Option Explicit

''
' Creates a cryptography object based on the name.
'
' @param AlgName The name of the cryptography object to be created.
' @return An instance of the object or Nothing if an empty name is provided.
' @remarks The supported names are:
' <pre>
' DES
' 3DES
' TripleDES
' Triple DES
' RC2
' Rijndael
' SHA
' SHA1
' HashAlgorithm
' SHA256
' SHA-256
' SHA384
' SHA-384
' SHA512
' SHA-512
' MD5
' HMACSHA1
' HMACSHA256
' HMACSHA384
' HMACSHA512
' HMACMD5
' HMACRIPEMD160
' MACTripleDES
' RSA
' DSA
' </pre>
' If a custom object is required, then the standard <Server>.<Class> format can be used.
'
Public Function CreateFromName(ByVal AlgName As String) As Object
    Dim Ret As Object
    
    Select Case LCase$(AlgName)
        Case "des":                             Set Ret = New DESCryptoServiceProvider
        Case "tripledes", "triple des", "3des": Set Ret = New TripleDESCryptoServiceProvider
        Case "rc2":                             Set Ret = New RC2CryptoServiceProvider
        Case "rijndael":                        Set Ret = New RijndaelManaged
        Case "sha", "sha1", "hashalgorithm":    Set Ret = New SHA1CryptoServiceProvider
        Case "sha256", "sha-256":               Set Ret = New SHA256Managed
        Case "sha512", "sha-512":               Set Ret = New SHA512Managed
        Case "sha384", "sha-384":               Set Ret = New SHA384Managed
        Case "md5":                             Set Ret = New MD5CryptoServiceProvider
        Case "hmacsha1":                        Set Ret = New HMACSHA1
        Case "hmacsha256":                      Set Ret = New HMACSHA256
        Case "hmacsha384":                      Set Ret = New HMACSHA384
        Case "hmacsha512":                      Set Ret = New HMACSHA512
        Case "hmacmd5":                         Set Ret = New HMACMD5
        Case "hmacripemd160":                   Set Ret = New HMACRIPEMD160
        Case "mactripledes":                    Set Ret = New MACTripleDES
        Case "rsa":                             Set Ret = New RSACryptoServiceProvider
        Case "dsa":                             Set Ret = New DSACryptoServiceProvider
        Case "":                                Set Ret = Nothing
        Case Else:                              Set Ret = CreateObject(AlgName)
    End Select
    
    Set CreateFromName = Ret
End Function

''
' Returns the encoded version of an Object Identifier.
'
' @param Str The object identifier as described in the X.690 standard.
' @return An encoded version of the object identifier as described in the X.690 standard.
'
Public Function EncodeOID(ByVal Str As String) As Byte()
    Const X680_OBJECT_IDENTIFIER_TYPE_TAG As Long = 6
    
    Dim EncodingStream  As New MemoryStream
    Dim Octets()        As String
    Octets = Split(Str, ".")
    
    If UBound(Octets) = 0 Then _
        Throw Cor.NewCryptographicException("Invalid OID.")
    
    On Error GoTo errTrap
    
    EncodingStream.WriteByte X680_OBJECT_IDENTIFIER_TYPE_TAG
    EncodingStream.WriteByte 0 ' place holder for the length.
    
    Dim Bytes() As Byte
    Dim Count   As Long
    Count = EncodeNumber(EncodeFirstTwoIdentifiers(Octets(0), Octets(1)), Bytes)
    EncodingStream.WriteBlock Bytes, 0, Count
    
    Dim i As Long
    For i = 2 To UBound(Octets)
        Count = EncodeNumber(Octets(i), Bytes)
        EncodingStream.WriteBlock Bytes, 0, Count
    Next i
    
    Dim Ret() As Byte
    Ret = EncodingStream.ToArray
    Ret(1) = EncodingStream.Length - 2 ' we don't count this byte or the tag byte.
    
    EncodeOID = Ret
    Exit Function
    
errTrap:
    Throw Cor.NewCryptographicException("Invalid OID.")
End Function

''
' Returns an object identifier for a specific name.
'
' @param Name The name of the object to obtain the identifier for.
' @return The object identifier (OID) for the object.
'
Public Function MapNameToOID(ByVal Name As String) As String
    Dim Ret As String
    
    Select Case LCase$(Name)
        Case "sha1", "sha": Ret = "1.3.14.3.2.26"
        Case "sha256":      Ret = "2.16.840.1.101.3.4.2.1"
        Case "sha384":      Ret = "2.16.840.1.101.3.4.2.2"
        Case "sha512":      Ret = "2.16.840.1.101.3.4.2.3"
        Case "md5":         Ret = "1.2.840.113549.2.5"
        Case "ripemd160":   Ret = "1.3.36.3.2.1"
        Case Else
            Ret = "1.2.840.113549.1.9.9"
    End Select

    MapNameToOID = Ret
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' X.690 8.19.2
Private Function EncodeNumber(ByVal Value As Long, ByRef Result() As Byte) As Long
    Dim Buf() As Byte
    ReDim Buf(15)
    
    Buf(0) = Value And &H7F
    Value = (Value And &HFFFFFF80) \ &H80
    
    Dim i As Long
    Do While Value > 0
        i = i + 1
        Buf(i) = (Value And &H7F) Or &H80
        Value = (Value And &HFFFFFF80) \ &H80
    Loop
    
    ReDim Preserve Buf(0 To i)
    
    Result = ReverseByteCopy(Buf)
    EncodeNumber = i + 1
End Function

' X.690 8.19.4
Private Function EncodeFirstTwoIdentifiers(ByVal x As Long, ByVal y As Long) As Long
    EncodeFirstTwoIdentifiers = (x * 40) + y
End Function

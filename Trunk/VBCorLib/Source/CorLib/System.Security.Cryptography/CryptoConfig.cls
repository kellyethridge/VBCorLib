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
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: CryptoConfig
'

''
' Used to access cryptography information.

Option Explicit

Private mDefaultAlgos As Hashtable

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
Public Function CreateFromName(ByRef AlgName As String) As IObject
    Set CreateFromName = CreateFromDefault(AlgName)
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
    
    On Error GoTo Catch
    
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
    
Catch:
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
'   Helpers
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

Private Function CreateFromDefault(ByRef AlgName As String) As IObject
    Dim Algo As Variant
    Algo = DefaultAlgos(AlgName)
    
    If Not IsEmpty(Algo) Then
        Set CreateFromDefault = CreateObject(Algo)
    End If
End Function

Private Property Get DefaultAlgos() As Hashtable
    Const DESCryptoServiceProviderType          As String = "CorLib.DESCryptoServiceProvider"
    Const TripleDESCryptoServiceProviderType    As String = "CorLib.TripleDESCryptoServiceProvider"
    Const RC2CryptoServiceProviderType          As String = "CorLib.RC2CryptoServiceProvider"
    Const RijndaelManagedType                   As String = "CorLib.RijndaelManaged"
    Const SHA1CryptoServiceProviderType         As String = "CorLib.SHA1CryptoServiceProvider"
    Const SHA256ManagedType                     As String = "CorLib.SHA256Managed"
    Const SHA512ManagedType                     As String = "CorLib.SHA512Managed"
    Const SHA384ManagedType                     As String = "CorLib.SHA384Managed"
    Const MD5CryptoServiceProviderType          As String = "CorLib.MD5CryptoServiceProvider"
    Const HMACSHA1Type                          As String = "CorLib.HMACSHA1"
    Const HMACSHA256Type                        As String = "CorLib.HMACSHA256"
    Const HMACSHA384Type                        As String = "CorLib.HMACSHA384"
    Const HMACSHA512Type                        As String = "CorLib.HMACSHA512"
    Const HMACMD5Type                           As String = "CorLib.HMACMD5"
    Const HMACRIPEMD160Type                     As String = "CorLib.HMACRIPEMD160"
    Const MACTripleDESType                      As String = "CorLib.MACTripleDES"
    Const RSACryptoServiceProviderType          As String = "CorLib.RSACryptoServiceProvider"
    Const DSACryptoServiceProviderType          As String = "CorLib.DSACryptoServiceProvider"
    Const RNGCryptoServiceProviderType          As String = "CorLib.RNGCryptoServiceProvider"
    
    If mDefaultAlgos Is Nothing Then
        Set mDefaultAlgos = Cor.NewHashtable(EqualityComparer:=StringComparer.OrdinalIgnoreCase)
        With mDefaultAlgos
            .Add "DES", DESCryptoServiceProviderType
            .Add "CorLib.DES", DESCryptoServiceProviderType
            .Add "TripleDES", TripleDESCryptoServiceProviderType
            .Add "Triple DES", TripleDESCryptoServiceProviderType
            .Add "3DES", TripleDESCryptoServiceProviderType
            .Add "RC2", RC2CryptoServiceProviderType
            .Add "CorLib.RC2", RC2CryptoServiceProviderType
            .Add "Rijndael", RijndaelManagedType
            .Add "CorLib.Rijndael", RijndaelManagedType
            .Add "CorLib.SymmetricAlgorithm", RijndaelManagedType
            .Add "SHA", SHA1CryptoServiceProviderType
            .Add "SHA1", SHA1CryptoServiceProviderType
            .Add "CorLib.SHA1", SHA1CryptoServiceProviderType
            .Add "CorLib.SHA1CryptoServiceProvider", SHA1CryptoServiceProviderType
            .Add "CorLib.HashAlgorithm", SHA1CryptoServiceProviderType
            .Add "SHA256", SHA256ManagedType
            .Add "SHA-256", SHA256ManagedType
            .Add "CorLib.SHA256", SHA256ManagedType
            .Add "SHA512", SHA512ManagedType
            .Add "SHA-512", SHA512ManagedType
            .Add "CorLib.SHA512", SHA512ManagedType
            .Add "SHA384", SHA384ManagedType
            .Add "SHA-384", SHA384ManagedType
            .Add "CorLib.SHA384", SHA384ManagedType
            .Add "MD5", MD5CryptoServiceProviderType
            .Add "CorLib.MD5", MD5CryptoServiceProviderType
            .Add "CorLib.MD5cryptoserviceprovider", MD5CryptoServiceProviderType
            .Add "HMACSHA1", HMACSHA1Type
            .Add "CorLib.HMAC", HMACSHA1Type
            .Add "CorLib.HMACSHA1", HMACSHA1Type
            .Add "CorLib.KeyedHashAlgorithm", HMACSHA1Type
            .Add "HMACSHA256", HMACSHA256Type
            .Add "CorLib.HMACSHA256", HMACSHA256Type
            .Add "HMACSHA384", HMACSHA384Type
            .Add "CorLib.HMACSHA384", HMACSHA384Type
            .Add "HMACSHA512", HMACSHA512Type
            .Add "CorLib.HMACSHA512", HMACSHA512Type
            .Add "HMACMD5", HMACMD5Type
            .Add "CorLib.HMACMD5", HMACMD5Type
            .Add "HMACRIPEMD160", HMACRIPEMD160Type
            .Add "CorLib.HMACRIPEMD160", HMACRIPEMD160Type
            .Add "MACTripleDES", MACTripleDESType
            .Add "CorLib.MACTripleDES", MACTripleDESType
            .Add "RSA", RSACryptoServiceProviderType
            .Add "CorLib.RSA", RSACryptoServiceProviderType
            .Add "CorLib.AsymmetricAlgorithm", RSACryptoServiceProviderType
            .Add "DSA", DSACryptoServiceProviderType
            .Add "CorLib.DSA", DSACryptoServiceProviderType
            .Add "RandomNumberGenerator", RNGCryptoServiceProviderType
            .Add "CorLib.RandomNumberGenerator", RNGCryptoServiceProviderType
            
            
        End With
    End If
    
    Set DefaultAlgos = mDefaultAlgos
End Property















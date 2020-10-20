VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2017 Kelly Ethridge
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
' Module: DSA
'

''
' Represents a base interface all DSA (Digital Signature Algorithm) implementations must implement.
'
' @see DSACryptoServiceProvider
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit


''
' Gets the name of the Key Exchange algorithm.
'
' @return The name of the algorithm.
'
Public Property Get KeyExchangeAlgorithm() As String: End Property

''
' Gets the key size for the current implementation.
'
' @return Returns the current key size in bits.
'
Public Property Get KeySize() As Long: End Property

''
' Sets the key size for the current implementation.
'
' @param Value The new key size in bits.
'
Public Property Let KeySize(ByVal Value As Long): End Property

''
' Gets a set of valid key sizes represented as KeySizes objects.
'
' @return An array of KeySizes objects.
'
Public Property Get LegalKeySizes() As KeySizes(): End Property

''
' Gets the name of the signature algorithm.
'
' @return The name of the signature algorithm.
'
Public Property Get SignatureAlgorithm() As String: End Property

''
' Releases any resources used by the current object.
'
Public Sub Clear(): End Sub

''
' Returns if the value and the object are equal.
'
' @param Value The value to check for equality.
' @return Returns True if the value and instance are equal, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Imports a DSA key from an XML string.
'
' @param XmlString The XML string containing the DSA key information.
'
Public Sub FromXmlString(ByRef XmlString As String): End Sub

''
' Returns a semi-unique number identifying this instance.
'
' @return A semi-unique number.
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string representation of this instance.
'
Public Function ToString() As String: End Function

''
' Exports the DSA key to an XML string.
'
' @param IncludePrivateParameters A flag indicating if the private portion of the key should be exported.
' @return An XML string of the DSA key.
'
Public Function ToXmlString(ByVal IncludePrivateParameters As Boolean) As String: End Function

''
' Creates a signature for a hash value.
'
' @param RgbHash The hash value to create a signature for.
' @return The signature created for the hash value.
'
Public Function CreateSignature(ByRef RgbHash() As Byte) As Byte(): End Function

''
' Exports the DSA key in a DSAParameters object.
'
' @param IncludePrivateParameters A flag indicating if the private portion of the key should be exported.
' @return A DSAParamters object containing the exported key.
'
Public Function ExportParameters(ByVal IncludePrivateParameters As Boolean) As DSAParameters: End Function

''
' Imports a DSA key from a DSAParameters object.
'
' @param Parameters The parameters object to import the key from.
'
Public Sub ImportParameters(ByVal Parameters As DSAParameters): End Sub

''
' Verifies a signature for a specified hash value.
'
' @param RgbHash The value the signature was created for.
' @param RgbSignature The signature to be verified.
' @return Returns True if the signature is for the hash value, False otherwise.
'
Public Function VerifySignature(ByRef RgbHash() As Byte, ByRef RgbSignature() As Byte) As Boolean: End Function


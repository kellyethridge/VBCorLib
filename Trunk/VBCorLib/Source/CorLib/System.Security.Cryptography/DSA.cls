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
'    Module: DSA
'

''
' Represents a base interface all DSA implementations must implement.
'
' @see DSACryptoServiceProvider
'
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
' @param RHS The new key size in bits.
'
Public Property Let KeySize(ByVal RHS As Long): End Property

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
Public Sub FromXmlString(ByVal XmlString As String): End Sub

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


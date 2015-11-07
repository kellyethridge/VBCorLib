VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RSA"
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
'    Module: RSA
'

''
' Represents the base interface from which all RSA class implementations must implement.
'
' @see RSACryptoServiceProvider
'
Option Explicit

''
' Returns the algorithm used for key exchange.
'
' @return The algorithm used for key exchange.
'
Public Property Get KeyExchangeAlgorithm() As String: End Property

''
' Returns the key size used by the RSA instance.
'
' @return The key size in bits used by the current instance.
'
Public Property Get KeySize() As Long: End Property

''
' Sets the key size used by the current RSA instance.
'
' @param RHS The new key size in bits.
'
Public Property Let KeySize(ByVal RHS As Long): End Property

''
' Returns an array of <b>KeySizes</b> objects representing all the valid key sizes.
'
' @return An array of <b>KeySizes</b> objects.
'
Public Property Get LegalKeySizes() As KeySizes(): End Property

''
' Returns the name of the signature algorithm used.
'
' @return The name of the signature algorithm.
'
Public Property Get SignatureAlgorithm() As String: End Property

''
' Releases any resources used by the RSA instance.
'
Public Sub Clear(): End Sub

''
' Returns if the Value is equal this RSA instance.
'
' @param Value The value to be tests for equality.
' @return Returns True if the value is equal to this instance, False otherwise.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Initializes an RSA instance with an Xml representation of an RSA key.
'
' @param XmlString The string to initialize the asymmetric algorithm with.
'
Public Sub FromXmlString(ByVal XmlString As String): End Sub

''
' Returns a hash code identifying this object instance.
'
' @return A hash code.
' @see IObject
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of this object instance.
'
' @return A string representation.
' @see IObject.
'
Public Function ToString() As String: End Function

''
' Returns an Xml representation of the key for this RSA instance.
'
' @param IncludePrivateParameters A flag indicating if the private portion of the key
' should be included in the Xml string.
' @return An Xml string representation of the RSA key.
' @remarks If the private portion is included, then another instance of an RSA implementation
' can use the key to encrypt values, otherwise, only enough information to decrypt is provided.
'
Public Function ToXmlString(ByVal IncludePrivateParameters As Boolean) As String: End Function

''
' Decrypts an array of encrypted data, returning the plain text data.
'
' @param Rgb The encrypted data to be decrypted.
' @return An array of plain text data.
'
Public Function DecryptValue(ByRef Rgb() As Byte) As Byte(): End Function

''
' Encrypts an array of plain text data, returning the encrypted data.
'
' @param Rgb The plain text data to be encrypted.
' @return The encrypted data.
'
Public Function EncryptValue(ByRef Rgb() As Byte) As Byte(): End Function

''
' Exports the RSA key in an <b>RSAParameters</b> object.
'
' @param IncludePrivateParameters A flag indicating if the private portion of the key should be exported.
' @return An <b>RSAParameters</b> object containing the RSA key.
' @remarks If the private portion is included, then another instance of an RSA implementation
' can use the key to encrypt values, otherwise, only enough information to decrypt is provided.
'
Public Function ExportParameters(ByVal IncludePrivateParameters As Boolean) As RSAParameters: End Function

''
' Imports the RSA key from an <b>RSAParameters</b> object.
'
' @param Parameters The object to import the RSA key from.
'
Public Sub ImportParameters(ByVal Parameters As RSAParameters): End Sub


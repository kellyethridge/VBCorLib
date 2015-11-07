VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsymmetricAlgorithm"
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
'    Module: AsymmetricAlgorithm
'

''
' Provides a common interface to be implemented by all Asymmetrical algorithm classes.
'
' @see RSACryptoServiceProvider
' @see DSACryptoServiceProvider
'
Option Explicit

''
' Returns the exchange algorithm used for key exchanges.
'
' @return A string name of the key exchange algorithm.
'
Public Property Get KeyExchangeAlgorithm() As String: End Property

''
' Returns the key size of the current algorithm instance.
'
' @return The key size being used by the current algorithm instance.
'
Public Property Get KeySize() As Long: End Property

''
' Sets the key size to be used by the current algorithm instance.
'
' @param RHS The new key size value.
'
Public Property Let KeySize(ByVal RHS As Long): End Property

''
' Returns an array of <b>KeySizes</b> representing all legal and valid key sizes
' for the current instance of the algorithm.
'
' @return A array of <b>KeySizes</b> objects.
'
Public Property Get LegalKeySizes() As KeySizes(): End Property

''
' Returns the name of the algorithm used for signing data and hashes.
'
' @return The name of the algorithm used for signing.
'
Public Property Get SignatureAlgorithm() As String: End Property

''
' Releases all resources used by the current algorithm instance.
'
Public Sub Clear(): End Sub

''
' Returns if the value and this object instance are equal.
'
' @param Value The value to test for equality.
' @return Returns True if the value and this object instance are equal, False otherwise.
'
Public Function Equals(ByRef Value As Variant) As Boolean: End Function

''
' Initializes an Asymmetric Algorithm instance from an XML string.
'
' @param XmlString The string to initialize the instance from.
'
Public Sub FromXmlString(ByVal XmlString As String): End Sub

''
' Returns a semi-unique value genetated for this instance.
'
' @return A semi-unique value.
'
Public Function GetHashCode() As Long: End Function

''
' Returns a string representation of this instance.
'
' @return A string describing this instance.
'
Public Function ToString() As String: End Function

''
' Returns an XML representation of the Asymmetrice Algorithm.
'
' @param IncludePrivateParameters A flag indicating if the private portion of the key should be included in the XML string.
' @return An XML string.
'
Public Function ToXmlString(ByVal IncludePrivateParameters As Boolean) As String: End Function


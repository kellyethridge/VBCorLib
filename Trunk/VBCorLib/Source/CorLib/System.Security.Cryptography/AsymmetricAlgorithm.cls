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
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: AsymmetricAlgorithm
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
' @param Value The new key size value.
'
Public Property Let KeySize(ByVal Value As Long): End Property

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
Public Sub FromXmlString(ByRef XmlString As String): End Sub

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


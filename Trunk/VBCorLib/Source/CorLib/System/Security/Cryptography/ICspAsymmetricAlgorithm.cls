VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICspAsymmetricAlgorithm"
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
'    Module: ICspAsymmetricAlgorithm
'

''
' Provices methods to allow access to Key container information and Crypto API blob import/export ability
' for Asymmetric algorithmes.
'
' @remarks This interface provides a common set of methods used to retrieve information about an Asymmetric
' algorithm's key container and provides the ability to import and export Byte arrays represetning Crypto
' API blobs used by Microsoft Cryptography Service Providers.
'
' @see RSACryptoServiceProvider
' @see DSACryptoServiceProvider
'
Option Explicit

''
' Returns a key container object holding information about an Asymmetric algorithm.
'
' @return A CspKeyContainerInfo object.
'
Public Property Get CspKeyContainerInfo() As CspKeyContainerInfo: End Property

''
' Returns a Crypto API blob from the Asymmetric algorithm.
'
' @param IncludePrivateParameters A flag indicating if the Blob should contain
' information considered private and sensative. The private information is usually
' information required to encrypt or sign data.
' @return A Byte array containing the Crypto API blob information.
'
Public Function ExportCspBlob(ByVal IncludePrivateParameters As Boolean) As Byte(): End Function

''
' Imports a Crypto API blob to be used by the Asymmetric algorithm.
'
' @param RawData The Byte array representing the Crypto API blob.
'
Public Sub ImportCspBlob(ByRef RawData() As Byte): End Sub



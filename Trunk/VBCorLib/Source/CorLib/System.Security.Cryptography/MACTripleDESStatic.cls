VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MACTripleDESStatic"
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
'    Module: MACTripleDESStatic
'

''
' Provides static methods associated with the MACTripleDES class.
'
' @see MACTripleDES
'
Option Explicit

''
' Returns an implementation of the MACTripleDES class.
'
' @param AlgName The name of the implementation to create and return.
' @return An implementation of a TripleDES MAC algorithm.
' @remarks By default the MACTripleDES implementation is returned. If a different
' implementation is desired, then use the standard 'Server.Class' naming convention.
'
Public Function Create(Optional ByVal AlgName As String = "MACTripleDES") As KeyedHashAlgorithm
    Set Create = CryptoConfig.CreateFromName(AlgName)
End Function

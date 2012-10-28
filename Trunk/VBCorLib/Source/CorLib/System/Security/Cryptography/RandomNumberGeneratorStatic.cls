VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RandomNumberGeneratorStatic"
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
'    Module: RandomNumberGeneratorStatic
'

''
' Provides static methods associated with the RandomNumberGenerator class.
'
Option Explicit


''
' Creates a new random number generator based on the name.
'
' @param RngName The name of the implementation to create.
' @return A new RandomNumberGenerator implementation.
' @remarks If the name is empty then RNGCryptoServiceProvider is returned.
'
Public Function Create(Optional ByVal RngName As String) As RandomNumberGenerator
    If Len(RngName) = 0 Then
        Set Create = New RNGCryptoServiceProvider
    Else
        Set Create = CreateObject(RngName)
    End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: ICollection
'

''
' Provides an interface to access size and enumeration capabilities of collections.
'
' @remarks<br>
' <p>Custom collection classes should implement this interface to allow
' VBCorLib to interact with it in way necessary. Method calls that add
' a range will usually accept objects implementing this interface.</p>
' <p>Classes implementing this interface should also implement the
' IEnumerable interface to maintain consistency with inherited interfaces.
' And, the author should consider implementing the NewEnum and GetEnumerator
' functions in the classes public interface.</p>
'
' @see IEnumerator
'
Option Explicit

''
' Returns an IEnumerator object.
'
' @return The enumerator.
'
Public Function GetEnumerator() As IEnumerator: End Function

''
' Returns an IEnumVariant object.
'
' @return The enumerator
'
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

''
' Returns the size of the collection.
'
' @return The size of the collection.
'
Public Property Get Count() As Long: End Property

''
' Copies the contents of the collection to an array.
'
' @param arr The array to copy the elements into.
' @param index The index in the destination array to start copying the elements to.
'
Public Sub CopyTo(ByRef arr As Variant, ByVal Index As Long): End Sub

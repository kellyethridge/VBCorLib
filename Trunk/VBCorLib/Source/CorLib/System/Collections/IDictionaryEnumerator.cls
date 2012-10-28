VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDictionaryEnumerator"
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
'    Module: IDictionaryEnumerator
'

''
' An interface used to enumerate a dictionary style collections (Hashtable.)
'
Option Explicit

Public Enum IDictionaryEnumType
    detEntries
    detKeys
    detValues
End Enum


''
' Returns a DictionaryEntry object for the current location in the collection.
'
' @return A DictionaryEntry containing key-value pairs.
'
Public Property Get Current() As Variant: End Property

''
' Moves the enumerator pointer to the next element in the collection.
'
' @return An indication of the success of the movement.
'
Public Function MoveNext() As Boolean: End Function

''
' Resets the enumerator to the beginning of the collection.
'
Public Sub Reset(): End Sub

''
' Returns a DictionaryEntry object for the current location in the collection.
'
' @return A DictionaryEntry containing key-value pairs.
'
Public Property Get Entry() As DictionaryEntry: End Property

''
' Returns the key portion of the current DictionaryEntry.
'
' @return The key of the current collection entry.
'
Public Property Get Key() As Variant: End Property

''
' Returns the value portion of the current DictionaryEntry.
'
' @return The value of the current collection entry.
'
Public Property Get Value() As Variant: End Property


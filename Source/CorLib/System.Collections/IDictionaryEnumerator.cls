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
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: IDictionaryEnumerator
'

''
' An interface used to enumerate a dictionary style collections (Hashtable.)
'
'@Folder("CorLib.System.Collections")
Option Explicit

Public Enum IDictionaryEnumType
    detEntries
    detKeys
    detValues
End Enum

Public Enum EnumerationType
    EnumEntries
    EnumKeys
    EnumValues
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


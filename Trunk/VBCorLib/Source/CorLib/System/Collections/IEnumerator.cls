VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEnumerator"
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
'    Module: IEnumerator
'

''
' This interface is used to allow for custom enumerators to be built.
'
' @remarks
' <p>The enumerator can be used manually to enumerate through the
' values of a list, or it can be wrapped in a custom enumerator to allow it to be
' used in the For..Each context.</p>
' <p>To use an IEnumerator object in For..Each, call the CreateEnumerator method
' passing in the IEnumerator object. The returned value is then returned in
' the standard NewEnum(-4) function called by For..Each. The enumeration will then
' be delegated to the custom IEnumerator object.<br><br>
' <b>Example</b><br>
' <pre>
' Public Function NewEnum() As IUnknown
'     Dim e As New MyEnumerator
'     e.Init Me
'     Set NewEnum = CreateEnumerator(e)
' End Function<br>
' </pre>
'
Option Explicit

''
' This returns the current item of the enumeration. The current
' item is called after a successful MoveNext.
'
' @return The current value in the enumeration.
'
Public Property Get Current() As Variant: End Property

''
' Requests the enumerator moves to the next item being enumerated.
' Returns True on success, False otherwise. This is called prior
' to Current. If this is successful, Current will be called.
'
' @return Indication if the move was successful.
'
Public Function MoveNext() As Boolean: End Function

''
' Requests the enumerator resets itself to begin enumerating from the beginning.
'
Public Sub Reset(): End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IComparable"
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
'    Module: IComparable
'

''
' Provides a standard interface used to compare an object to a value.
'
' @remarks
' This interface is required to allow collections to compare two objects
' when sorting or searching. If there are objects in the collection that don't
' implement this interface, then those methods will fail unless an custom
' comparer is supplied.
'
Option Explicit

''
' Compares a value to the current object instance.
'
' @param value The value to compare this object to.
' @return Value indicating its relationship to the value.
' @remarks Comparison guide:<br>
' This object is the same value as value returns a 0.<br>
' This object is less than the value returns a negative value.<br>
' This object is greater than the value returns a positive value.<br>
'
Public Function CompareTo(ByRef value As Variant) As Long: End Function


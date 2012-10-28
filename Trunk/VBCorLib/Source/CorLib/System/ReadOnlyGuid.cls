VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyGuid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: ReadOnlyGuid
'
Option Explicit
Implements IObject
Implements Guid
Implements IComparable
Implements IFormattable


Private mGuid As Guid


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal g As Guid)
    Set mGuid = g
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = mGuid.Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = mGuid.GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = mGuid.ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Guid Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Guid_CompareTo(Value As Variant) As Long
    Guid_CompareTo = mGuid.CompareTo(Value)
End Function

Private Function Guid_Equals(Value As Variant) As Boolean
    Guid_Equals = mGuid.Equals(Value)
End Function

Private Function Guid_GetHashCode() As Long
    Guid_GetHashCode = mGuid.GetHashCode
End Function

Private Property Get Guid_Handle() As Long
    Throw Cor.NewNotSupportedException("Guid is read-only.")
End Property

Private Property Get Guid_IsReadOnly() As Boolean
    Guid_IsReadOnly = True
End Property

Private Function Guid_ToByteArray() As Byte()
    Guid_ToByteArray = mGuid.ToByteArray
End Function

Private Function Guid_ToString(Optional ByVal Format As String, Optional ByVal provider As IFormatProvider) As String
    Guid_ToString = mGuid.ToString(Format, provider)
End Function

Private Function Guid_ToProgramID() As String
    Guid_ToProgramID = mGuid.ToProgramID
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = mGuid.CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByVal Format As String, ByVal provider As IFormatProvider) As String
    IFormattable_ToString = mGuid.ToString(Format, provider)
End Function

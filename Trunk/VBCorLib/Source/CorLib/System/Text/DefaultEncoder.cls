VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultEncoder"
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
'    Module: DefaultEncoder
'
Option Explicit
Implements IObject
Implements Encoder

Private mEncoding As Encoding



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Encoding As Encoding)
    Set mEncoding = Encoding
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function

Private Function IObject_Equals(ByRef Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Encoder Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Encoder_Equals(Value As Variant) As Boolean
    Encoder_Equals = IObject_Equals(Value)
End Function

Private Function Encoder_GetByteCount(Chars() As Integer, ByVal Index As Long, ByVal Count As Long, ByVal Flush As Boolean) As Long
    Encoder_GetByteCount = mEncoding.GetByteCount(Chars, Index, Count)
End Function

Private Function Encoder_GetBytes(Chars() As Integer, ByVal CharIndex As Long, ByVal CharCount As Long, Bytes() As Byte, ByVal ByteIndex As Long, ByVal Flush As Boolean) As Long
    Encoder_GetBytes = mEncoding.GetBytesEx(Chars, CharIndex, CharCount, Bytes, ByteIndex)
End Function

Private Function Encoder_GetHashCode() As Long
    Encoder_GetHashCode = IObject_GetHashCode
End Function

Private Function Encoder_ToString() As String
    Encoder_ToString = IObject_ToString
End Function

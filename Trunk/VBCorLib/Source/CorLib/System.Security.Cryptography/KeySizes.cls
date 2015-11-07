VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeySizes"
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
'    Module: KeySizes
'

''
' Represents a set of key sizes.
'
' @remarks All of the key sizes are represented by indicating a minimum key size and a maximum key size,
' then using a step to fill in key sizes between the two.
'
Option Explicit
Implements IObject

Private mMinSize    As Long
Private mMaxSize    As Long
Private mSkipSize   As Long



''
' The mininum key size in bits.
'
' @return The mininum key size in bits.
'
Public Property Get MinSize() As Long
    MinSize = mMinSize
End Property

''
' The maximum key size in bits.
'
' @return The maximum key size in bits.
'
Public Property Get MaxSize() As Long
    MaxSize = mMaxSize
End Property

''
' The step size between keys in bits.
'
' @returns the step size in bits.'
'
Public Property Get SkipSize() As Long
    SkipSize = mSkipSize
End Property

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MinSize As Long, ByVal MaxSize As Long, ByVal SkipSize As Long)
    mMinSize = MinSize
    mMaxSize = MaxSize
    mSkipSize = SkipSize
End Sub

Friend Function IsLegalSize(ByVal BitLength As Long) As Boolean
    If mSkipSize = 0 Then
        IsLegalSize = (BitLength = mMinSize)
        Exit Function
    End If
    
    Dim i As Long
    For i = mMinSize To mMaxSize Step mSkipSize
        If BitLength = i Then
            IsLegalSize = True
            Exit Function
        End If
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaseInsensitiveHashCodePrvdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: CaseInsensitiveHashCodePrvdr
'

''
' Generates hashcodes for values as normal, but ignores the case for strings.
'
Option Explicit
Implements IObject
Implements IHashcodeProvider

Private mStringHashChars As WordBuffer


''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @param Value The value or variable to generate a hashcode for.
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode(Optional ByRef Value As Variant) As Long
    If Not IsMissing(Value) Then
        GetHashCode = IHashcodeProvider_GetHashCode(Value)
    Else
        GetHashCode = ObjPtr(CUnk(Me))
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashString(ByRef Key As Variant) As Long
    Dim StringLength As Long
    StringLength = Len(Key)
    
    With mStringHashChars.SA
        .cElements = StringLength
        .pvData = StrPtr(Key)
    End With

    Dim i   As Long
    Dim Ch  As Long
    For i = 0 To StringLength - 1
        Ch = CharLower(mStringHashChars.Data(i))
        HashString = ((HashString * 16) + Ch) And &H3FFFFFF
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IHashCodeProvider Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IHashcodeProvider_GetHashCode(Value As Variant) As Long
    If VarType(Value) = vbString Then
        IHashcodeProvider_GetHashCode = HashString(Value)
    Else
        IHashcodeProvider_GetHashCode = Object.GetHashCode(Value)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWordBuffer(mStringHashChars, 0, &H7FFFFFFF)
End Sub

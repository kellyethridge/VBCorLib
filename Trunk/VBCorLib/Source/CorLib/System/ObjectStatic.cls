VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: ObjectStatic
'

''
' Provides default implementations of the IObject interface.
'
Option Explicit

Private mStringHashChars As WordBuffer


''
' Returns a default hashcode for the specific value.
'
' @param Value The value to generate the hashcode for.
' @return The hashcode value.
' @remarks An object should never pass itself into this
' function. Instead, the object should implement the default
' hashcode by using ObjPtr.
' <pre>
' Public Function GetHashCode As Long
'     GetHashCode = ObjPtr(CUnk(Me))
' End Function
' </pre>
'
Public Function GetHashCode(ByRef Value As Variant) As Long
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbBoolean
            GetHashCode = CLng(Value)
            
        Case vbString
            GetHashCode = HashString(Value)
        
        Case vbDouble
            With AsDLong(CDbl(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbSingle
            GetHashCode = AsLong(CSng(Value))
            
        Case vbDate
            With AsDLong(CDate(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbCurrency
            With AsDLong(CCur(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbObject, vbDataObject
            If Value Is Nothing Then
                Exit Function
            End If
            
            If TypeOf Value Is IObject Then
                Dim Obj As IObject
                Set Obj = Value
                GetHashCode = Obj.GetHashCode
            Else
                GetHashCode = ObjPtr(CUnk(Value))
            End If
    End Select
End Function

''
' Determines if the value and the object are the same instance.
'
' @param YourObject The object to compare the value against.
' @param Value The value to compare against the object.
' @return Returns True if the value and the object are the same instance, False otherwise.
'
Public Function Equals(ByVal YourObject As Object, ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        Equals = (Value Is YourObject)
    End If
End Function

''
' Returns a default string representation of an object.
'
' @param YourObject The object to derive the name from.
' @param YourApp the App object from the calling function.
' @return A string representation of the object.
'
Public Function ToString(ByVal YourObject As Object, ByVal YourApp As Object) As String
    If YourObject Is Nothing Then Exit Function
    
    ToString = TypeName(YourObject)
    If Not YourApp Is Nothing Then
        Dim AppObject As App
        Set AppObject = YourApp
        ToString = AppObject.Title & "." & ToString
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashString(ByRef Value As Variant) As Long
    Dim l As Long
    l = Len(Value)
    With mStringHashChars.SA
        .cElements = l
        .pvData = StrPtr(Value)
    End With
    
    Dim i As Long
    For i = 0 To l - 1
       HashString = ((HashString * 16) + mStringHashChars.Data(i)) And &H3FFFFFF
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWordBuffer(mStringHashChars, 0, 0)
End Sub

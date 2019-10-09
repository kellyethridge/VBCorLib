VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectBase"
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
' Module: ObjectBase
'

''
' This class represent a proxy as a base class for those classes that implement <a href="IObject.html">IObject</a>.
' It is intended to give default behavior for the three main methods of the Object class in .NET.
' The base class is accessed using the <a href="StaticClasses.Get.MyBase.html">MyBase</a> global static property.
'
' This class is not intended to be used by code other than in the methods of the <a href="IObject.html">IObject</a>
' interface implementations.
'
' @include "..\DocIncludes\ObjectBase.html"
Option Explicit


''
' Default object comparison. Simply checks for reference to same instance.
'
' @param ObjA The subclass requesting the call.
' @param ObjB The value to be equating the subclass to.
' @return <b>True</b> if both values point to the same object instance; <b>False</b> otherwise.
' @include "..\DocIncludes\ObjectBase.html"
Public Function Equals(ByVal ObjA As IObject, ByRef ObjB As Variant) As Boolean
    If IsObject(ObjB) Then
        Equals = ObjA Is ObjB
    End If
End Function

''
' Returns a hashcode for the specified object instance.
'
' @param Obj The object to generate the hashcode for.
' @return Returns a Long as the representation of the hashcode.
' @include "..\DocIncludes\ObjectBase.html"
Public Function GetHashCode(ByVal Obj As IObject) As Long
    GetHashCode = ObjPtr(CUnk(Obj))
End Function

''
' Returns a string representation of the object instance with optional project name.
'
' @param Obj The object to return a string representation of.
' @param AppObject The classes App object used to determine the project name.
' @return Returns a combination of the object type name and project name.
' @include "..\DocIncludes\ObjectBase.html"
Public Function ToString(ByVal Obj As IObject, Optional ByVal AppObject As Object) As String
    If Obj Is Nothing Then
        Exit Function
    End If
    
    ToString = TypeName(Obj)
    
    If Not AppObject Is Nothing Then
        If TypeOf AppObject Is App Then
            Dim OtherApp As App
            Set OtherApp = AppObject
            ToString = OtherApp.Title & "." & ToString
        End If
    End If
End Function

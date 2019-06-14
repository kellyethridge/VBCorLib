VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: WeakReference
'

''
' Provides a reference to an object, but does not keep that object alive.
'
' @remarks
' This is to help with circular references to prevent unreachable objects from
' remaining in memory. The number of references is examined after each call to
' an object's IUnknown.Release method. Once it reaches zero, the WeakReference
' object clears the hook into the referenced object and sets the IsAlive status to false.
'
' @see Constructors
'
Option Explicit
Implements IObject

''
' This event is raised when a controlling WeakReference loses
' reference to the underlying Target object.
'
Public Event Disposed()


Private mPtrWeak                As Long



''
' Returns a strong reference to the object a weak reference is being kept to.
'
' @return Strong referenced object.
'
Public Property Get Target() As Object
    If Not IsAlive Then _
        Throw Cor.NewInvalidOperationException("Object reference has already been terminated.")
    
    Dim Ret As IUnknown
    ObjectPtr(Ret) = MemLong(mPtrWeak + 20)
    Set Target = Ret
    ObjectPtr(Ret) = 0
End Property

''
' Sets the target object to keep a weak reference to.
'
' @param RHS Strong object to create a weak reference to.
'
Public Property Set Target(ByVal RHS As Object)
    Call Init(RHS)
End Property

''
' Returns the status of the weakly referenced object.
'
' @return Status of the weak referenced object.
'
Public Property Get IsAlive() As Boolean
    IsAlive = mPtrWeak <> vbNullPtr
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Target As Object)
    If Target Is Nothing Then Exit Sub
    mPtrWeak = InitWeakReference(Me, Target)
End Sub


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

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
'    Module: WeakReference
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
Private mIsChild                As Boolean
Private mChildWeakReferences    As New Collection



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
    If IsChild Then
        If Not RHS Is Nothing Then Throw Cor.NewInvalidOperationException("Can only set Target to Nothing on a child WeakReference.")
    End If
    
    Call Dispose
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

''
' Creates a new weak reference without allocating addition
' resources to maintain reference integrity.
'
' @return A child version of the parent weak reference.
' @remarks If multiple weak references want to be maintained
' on a single object, creating child weak references is the
' most efficient method.
' <p>You cannot set the target to anything other than Nothing
' on a child weak reference. This is to preserve the integerity
' of the controlling weak reference and the referenced object.</p>
' <p>If the controlling weak reference loses reference to the
' Target object, all child weak references will also lose reference.</p>
'
Public Function CreateChild() As WeakReference
    If IsChild Then _
        Throw Cor.NewInvalidOperationException("Child reference cannot create child reference.")
    
    Set CreateChild = New WeakReference
    Call mChildWeakReferences.Add(CreateChild, CStr(CreateChild.GetHashCode))
    Call CreateChild.InitChild(mPtrWeak)
End Function

''
' Returns if this isnstance of WeakReference is a child.
'
' @return Returns True if this instance was created using
' <i>Object.CreateChild</i>, otherwise False.
'
Public Property Get IsChild() As Boolean
    IsChild = mIsChild
End Property

''
' Releases the reference to the Target object.
'
' @param ChildrenOnly True is this instance is the controlling
' WeakReference and only the children are to be disposed of.
' @remarks This allows the weak reference and its children to be
' disconnected from the Target object. The children themselves can
' be disposed of without disposing the controlling WeakReference.
'
Public Sub Dispose(Optional ByVal ChildrenOnly As Boolean = False)
    If Not IsAlive Then Exit Sub
    
    ' Are we the controlling WeakReference and are
    ' disposing everything, not just our children?
    If (Not IsChild) And (Not ChildrenOnly) Then
        ' We are the controlling WeakReference, so
        ' Set the ".pOwner" to Nothing. This will
        ' signal the Release method that we are no
        ' longer maintaining a weak reference.
        MemLong(mPtrWeak + 24) = vbNullPtr
        
        ' And now set our flag to "Dead"
        mPtrWeak = vbNullPtr
    ElseIf IsChild Then
        ' Children always dispose of there weak reference.
        mPtrWeak = vbNullPtr
    End If
    
    ' Only the controller will have a collection of kids.
    Dim Child As WeakReference
    For Each Child In mChildWeakReferences
        Call Child.Dispose
    Next Child
    
    ' Now that they are all disposed, get rid of them.
    Set mChildWeakReferences = New Collection
    
    ' All WeakRefernece objects, including the children
    ' will raise an event on disposing.
    RaiseEvent Disposed
End Sub

''
' Releases any dead child WeakReferences.
'
' @remarks This releases any children that may have been disposed
' by calling Dispose on the child WeakReference and not through
' the Dispose through the controlling WeakReference object.
'
Public Sub ClearDeadReferences()
    Dim Child As WeakReference
    Dim NewChildren As Collection
    Set NewChildren = New Collection
    
    For Each Child In mChildWeakReferences
        If Child.IsAlive Then Call NewChildren.Add(Child, CStr(Child.GetHashCode))
    Next Child
    
    Set mChildWeakReferences = NewChildren
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Target As Object)
    If Target Is Nothing Then Exit Sub
    mPtrWeak = InitWeakReference(Me, Target)
End Sub

Friend Sub InitChild(ByVal ptr As Long)
    mPtrWeak = ptr
    mIsChild = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Call Dispose
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

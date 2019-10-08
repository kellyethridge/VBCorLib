Attribute VB_Name = "Delegation"
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
' Module: Delegation
'
Option Explicit

Private Const DelegationCode        As Currency = -368956918007638.6215@     ' delegator code from Matt Curland
Private Const SizeOfLocalDelegate   As Long = 12

Private Type LocalDelegateVTable
    Func(3) As Long
End Type

Private Type DelegateVTable
    Func(7) As Long
End Type

Private Type LocalDelegate
    pVTable     As Long
    pfn         As Long
    cRefs       As Long
End Type

Public Type Delegate
    pVTable As Long
    pfn     As Long
End Type

Private mDelegateCode       As Currency
Private mDelegateTemplate   As LocalDelegate
Private mLocalVTable        As LocalDelegateVTable
Private mLocalVTablePtr     As Long
Private mVTable             As DelegateVTable
Private mOkVTablePtr        As Long
Private mFailVTablePtr      As Long
Private mFunc_T_T_LongDel   As Delegate
Private mFunc_T_T_T_StringDel As Delegate


''
' This creates a new lightweight function delegator object
' that calls a user specified function using the AddressOf operator.
'
' @param pfn The address to function to be called.
' @return A lightweight COM object used to call a function.
' @remarks When using this method of delegate creation by a class that will live for the
' duration of an application, the class should unhook the reference using ObjectPtr instead
' of letting the reference be set to Nothing and causing a call to the Release method in this
' module. This is to prevent a possible call to the Release method after this module has been
' deallocated during the tear-down phase of an application, since the delegate may be deallocated
' after this module.
Public Function NewDelegate(ByVal pfn As Long) As IUnknown
    Dim This As Long
    
    Init
    This = CoTaskMemAlloc(SizeOfLocalDelegate)
    
    If This = vbNullPtr Then _
        Throw New OutOfMemoryException
    
    mDelegateTemplate.pfn = pfn
    mDelegateTemplate.pVTable = mLocalVTablePtr
    CopyMemory ByVal This, mDelegateTemplate, SizeOfLocalDelegate
    ObjectPtr(NewDelegate) = This
End Function

''
' Initializes an existing delegate light-weight object with an optional function pointer.
'
' @param Struct The delegate object to be initialized.
' @param pfn The initial function pointer the delegate will be set to.
' @return An object reference to the delegate passed in.
' @remarks When using this method of delegate creation by a class that will live for the
' duration of an application, the class should unhook the reference using ObjectPtr instead
' of letting the reference be set to Nothing and causing a call to the Release method in this
' module. This is to prevent a possible call to the Release method after this module has been
' deallocated during the tear-down phase of an application, since the delegate may be deallocated
' after this module.
Public Function InitDelegate(ByRef Struct As Delegate, Optional ByVal pfn As Long) As IUnknown
    Init
    Struct.pfn = pfn
    Struct.pVTable = mOkVTablePtr
    ObjectPtr(InitDelegate) = VarPtr(Struct)
End Function

Public Function CallFunc_T_T_Long(ByVal lpFunc As Long, ByVal lpArg1 As Long, ByVal lpArg2 As Long) As Long
    Dim Caller As Func_T_T_Long
    
    Init
    mFunc_T_T_LongDel.pfn = lpFunc
    ObjectPtr(Caller) = VarPtr(mFunc_T_T_LongDel)
    
    On Error GoTo Catch
    CallFunc_T_T_Long = Caller.Invoke(ByVal lpArg1, ByVal lpArg2)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    ObjectPtr(Caller) = vbNullPtr
    Return
End Function

Public Function CallFunc_T_T_T_String(ByVal lpFunc As Long, ByVal lpArg1 As Long, ByVal lpArg2 As Long, ByVal lpArg3 As Long) As String
    Dim Caller As Func_T_T_T_String
    
    Init
    mFunc_T_T_T_StringDel.pfn = lpFunc
    ObjectPtr(Caller) = VarPtr(mFunc_T_T_T_StringDel)
    
    On Error GoTo Catch
    CallFunc_T_T_T_String = Caller.Invoke(ByVal lpArg1, ByVal lpArg2, ByVal lpArg3)
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    ObjectPtr(Caller) = vbNullPtr
    Return
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If mLocalVTablePtr = vbNullPtr Then
        mDelegateCode = DelegationCode
        VirtualProtect mDelegateCode, 8, PAGE_EXECUTE_READWRITE, 0&
        
        With mLocalVTable
            .Func(0) = FuncAddr(AddressOf LocalDelegate_QueryInterface)
            .Func(1) = FuncAddr(AddressOf LocalDelegate_AddRef)
            .Func(2) = FuncAddr(AddressOf LocalDelegate_Release)
            .Func(3) = VarPtr(mDelegateCode)
        End With
        
        mLocalVTablePtr = VarPtr(mLocalVTable)
        mDelegateTemplate.cRefs = 1
        
        With mVTable
            .Func(0) = FuncAddr(AddressOf Delegate_OKQueryInterface)
            .Func(1) = FuncAddr(AddressOf Delegate_AddRefRelease)
            .Func(2) = .Func(1)
            .Func(3) = VarPtr(mDelegateCode)
            .Func(4) = FuncAddr(AddressOf Delegate_FailQueryInterface)
            .Func(5) = .Func(1)
            .Func(6) = .Func(1)
            .Func(7) = VarPtr(mDelegateCode)
            
            mOkVTablePtr = VarPtr(.Func(0))
            mFailVTablePtr = VarPtr(.Func(4))
        End With
        
        mFunc_T_T_LongDel.pVTable = mOkVTablePtr
        mFunc_T_T_T_StringDel.pVTable = mOkVTablePtr
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   VTable functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LocalDelegate_QueryInterface(ByRef This As LocalDelegate, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.cRefs = This.cRefs + 1
End Function

Private Function LocalDelegate_AddRef(ByRef This As LocalDelegate) As Long
    This.cRefs = This.cRefs + 1
    LocalDelegate_AddRef = This.cRefs
End Function

Private Function LocalDelegate_Release(ByRef This As LocalDelegate) As Long
    This.cRefs = This.cRefs - 1
    
    If This.cRefs = 0 Then
        CoTaskMemFree VarPtr(This)
    Else
        LocalDelegate_Release = This.cRefs
    End If
End Function

Private Function Delegate_OKQueryInterface(ByRef This As Delegate, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = VarPtr(This)
    This.pVTable = mFailVTablePtr
End Function

Private Function Delegate_FailQueryInterface(ByRef This As Delegate, ByVal riid As Long, ByRef pvObj As Long) As Long
    pvObj = vbNullPtr
    Delegate_FailQueryInterface = E_NOINTERFACE
End Function

Private Function Delegate_AddRefRelease(ByVal This As Long) As Long
    ' do nothing
End Function



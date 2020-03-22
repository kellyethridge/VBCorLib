VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureAwareStringComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The MIT License (MIT)
'Copyright (c) 2020 Kelly Ethridge
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
' Module: CultureAwareStringComparer
'

Option Explicit
Implements IObject
Implements StringComparer
Implements IComparer
Implements IEqualityComparer

Private mCulture As CultureInfo
Private mOptions As StringComparison



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Culture As CultureInfo, ByVal IgnoreCase As Boolean)
    Set mCulture = Culture
    mOptions = IIfLong(IgnoreCase, CompareOptions.IgnoreCase, CompareOptions.None)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   StringComaprer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function StringComparer_Compare(x As Variant, y As Variant) As Long
    If VarType(x) = vbString Then
        If VarType(y) = vbString Then
            StringComparer_Compare = CorString.CompareCulturalV(x, y, mCulture, mOptions)
            Exit Function
        End If
    End If
    
    StringComparer_Compare = Comparer.Default.Compare(x, y)
End Function

Private Function StringComparer_Equals(x As Variant, y As Variant) As Boolean
    If VarType(x) = vbString Then
        If VarType(y) = vbString Then
            StringComparer_Equals = CorString.CompareCulturalV(x, y, mCulture, mOptions) = 0
            Exit Function
        End If
    End If
    
    StringComparer_Equals = Object.Equals(x, y)
End Function

Private Function StringComparer_GetHashCode(Obj As Variant) As Long
    If VarType(Obj) = vbString Then
        If mOptions = CompareOptions.IgnoreCase Then
            StringComparer_GetHashCode = Object.GetHashCode(UCase$(Obj))
            Exit Function
        End If
    End If
    
    StringComparer_GetHashCode = Object.GetHashCode(Obj)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(a As Variant, b As Variant) As Long
    IComparer_Compare = StringComparer_Compare(a, b)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEqualityComparer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEqualityComparer_Equals(x As Variant, y As Variant) As Boolean
    IEqualityComparer_Equals = StringComparer_Equals(x, y)
End Function

Private Function IEqualityComparer_GetHashCode(Obj As Variant) As Long
    IEqualityComparer_GetHashCode = StringComparer_GetHashCode(Obj)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = MyBase.Equals(Me, Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = MyBase.GetHashCode(Me)
End Function

Private Function IObject_ToString() As String
    IObject_ToString = MyBase.ToString(Me, App)
End Function


Attribute VB_Name = "Interaction"
'The MIT License (MIT)
'Copyright (c) 2016 Kelly Ethridge
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
' Module: CorInteraction
'
Option Explicit

Public Function IfObject(ByVal ObjA As Object, ByVal ObjB As Object) As Object
    If ObjA Is Nothing Then
        Set IfObject = ObjB
    Else
        Set IfObject = ObjA
    End If
End Function

Public Function IfString(ByRef a As String, ByRef b As String) As String
    If LenB(a) > 0 Then
        IfString = a
    Else
        IfString = b
    End If
End Function

Public Function IIfLong(ByVal Expression As Boolean, ByVal TruePart As Long, ByVal FalsePart As Long) As Long
    If Expression Then
        IIfLong = TruePart
    Else
        IIfLong = FalsePart
    End If
End Function

Public Function IsPicture(ByRef Value As Variant) As Boolean
    Dim Pic As IPicture
    
    On Error GoTo Catch
    Set Pic = Value
    IsPicture = Not Pic Is Nothing
    
Catch:
End Function

Public Function IsPictureResourceGroup(ByRef Value As Variant) As Boolean
    Dim Group As PictureResourceGroup
    
    On Error GoTo Catch
    Set Group = Value
    IsPictureResourceGroup = Not Group Is Nothing
    
Catch:
End Function

Public Function IsInteger(ByRef Value As Variant) As Boolean
    ' we use VBVM6.VariantType because VarType will cause objects
    ' that have a default property to return the default value.
    ' If the default property happens to return a numeric value
    ' it will give a false positive here.
    Select Case CorVarType(Value)
        Case vbLong, vbInteger, vbByte
            IsInteger = True
    End Select
End Function

Public Function CorVarType(ByRef Value As Variant) As VbVarType
    CorVarType = VariantType(Value) And &HFF ' we mask because VariantType will include BY_REF (&h4000) flag if it exists.
End Function

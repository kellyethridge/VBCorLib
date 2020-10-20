VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2014 Kelly Ethridge
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
' Module: Comparer
'

''
' Compares two values for equivalence, where string comparisons are case-sensitive.
'
' @remarks This class is the default implementation of the <a href="IComparer.html">IComparer</a> interface. The <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> class
' is the implementation of the <a href="IComparer.html">IComparer</a> interface that performs case-insensitive string comparisons.
' <p>Comparison procedures use the <a href="ThreadStatic.Get.CurrentCulture.html">Thread.CurrentCulture</a> of the current thread unless otherwise specified. String comparisons might have different results depending on the culture.</p>
' @see ComparerStatic
' @see Constructors
' @see IComparer
' @see IComparable
' @see CaseInsensitiveComparer
' @see ThreadStatic
' @include "..\DocIncludes\Comparer.html"
'@Folder("CorLib.System.Collections")
Option Explicit
Implements IObject
Implements IComparer

Private mCulture As CultureInfo


''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
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
' Returns a string representation of this object instance.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Performs a case-sensitive comparison of two values of the same type and returns a value indicating whether one is less than, equal to, or greater than the other.
'
' @param a The left-hand-side of the equation.
' @param b the right-hand-side of the equation.
' @return A signed integer that indicates the relative values of <i>a</i> and <i>b</i>, as shown in the following table.
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Meaning</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td><i>a</i> is less than <i>b</i>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td><i>a</i> equals <i>b</i>.</td>
'    </tr>
'    <tr>
'        <td>Greater than zero</td><td><i>a</i> is greater than <i>b</i>.</td>
'    </tr>
' </table>
' @remarks If <i>a</i> implements <a href="IComparable.html">IComparable</a>, then <i>a</i>.<b>CompareTo</b>(<i>b</i>)
' is returned; otherwise, if <i>b</i> implements <a href="IComparable.html">IComparable</a>, then the negated result of <i>b</i>.<b>CompareTo</b>(<i>a</i>) is returned.
' <p>String comparisons might have different results depending on the culture.</i>
' @see IComparable
' @see CultureInfo
' @include "..\DocIncludes\Comparer.Compare.html"
Public Function Compare(ByRef a As Variant, ByRef b As Variant) As Long
    Dim Comparable As IComparable
    Dim AVarType As VbVarType
    Dim BVarType As VbVarType
    
    AVarType = VarType(a)
    BVarType = VarType(b)
    
    Select Case AVarType
        Case vbString
            If BVarType = vbString Then
                Compare = CorString.CompareCulturalV(a, b, mCulture, CompareOptions.None)
                Exit Function
            End If
        Case vbNull
            If BVarType <> vbNull Then
                Compare = -1
            End If
            Exit Function

        Case vbEmpty
            If BVarType = vbNull Then
                Compare = 1
            ElseIf BVarType <> vbEmpty Then
                Compare = -1
            End If
            Exit Function
            
        Case vbObject, vbDataObject
            If Not a Is Nothing Then
                If TypeOf a Is IComparable Then
                    Set Comparable = a
                    Compare = Comparable.CompareTo(b)
                    Exit Function
                ElseIf a Is b Then
                    Compare = 0
                    Exit Function
                End If
            End If
        
        Case vbUserDefinedType
            If BVarType = vbUserDefinedType Then
                If IsInt64(a) And IsInt64(b) Then
                    Compare = CallFunc_T_T_Long(AddressOf Int64Compare, DataPtr(a), DataPtr(b))
                    Exit Function
                End If
            End If
        
        Case BVarType
            If a < b Then
                Compare = -1
            ElseIf a > b Then
                Compare = 1
            End If
            Exit Function
    End Select

    Select Case BVarType
        Case vbNull, vbEmpty
            Compare = 1
            
        Case vbObject, vbDataObject
            If b Is Nothing Then
                Exit Function
            ElseIf TypeOf b Is IComparable Then
                Set Comparable = b
                Compare = -Comparable.CompareTo(a)
            Else
                Error.Argument Argument_ImplementIComparable
            End If
            
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_InvalidValueType, TypeName(a)))
    End Select
End Function

Friend Property Get LCID() As Long
    LCID = mCulture.LCID
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Culture As CultureInfo)
    If Culture Is Nothing Then _
        Error.ArgumentNull "Culture"
        
    Set mCulture = Culture
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparer_Compare(a As Variant, b As Variant) As Long
    IComparer_Compare = Compare(a, b)
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

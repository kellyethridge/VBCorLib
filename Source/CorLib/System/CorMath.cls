VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorMath"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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
' Module: CorMath
'

''
' Provides extended math functionality.
'
' @remarks The functions in this class are global. To access them simply call
' the functions directly as any normal function. The CorMath class does not need
' to be referenced.
'
' <pre>
' Debug.Print PI
' Debug.Print Max(1, 2)
' </pre>
'
'@Folder("CorLib.System")
Option Explicit

Private Const corPI         As Double = 3.14159265358979
Private Const corAtn1       As Double = 0.785398163397448
Private Const corLog10      As Double = 2.30258509299405
Private Const corNegZero    As Currency = -922337203685477.5807@ - 0.0001@   ' VB won't accept -922337203685477.5808@ as a constant

Private mNegativeZero As Double


''
' Returns the next highest whole integer value.
'
' @param Value The number to round to the next integer.
' @return The next integer.
' @remarks Negative values will round towards zero. If the
' value is already a whole integer value then that value will be returned.
'
Public Function Ceiling(ByVal Value As Double) As Double
    If Value < 0# Then
        Ceiling = Fix(Value)
    Else
        Ceiling = -Int(-Value)
    End If
End Function

''
' Returns the next lowest whole integer value.
'
' @param Value the number to be rounded downward.
' @return The next lowest whole integer value.
' @remarks Negative values will be rounded away from zero. -2.1 will be
' rounded down to -3.
'
Public Function Floor(ByVal Value As Double) As Double
    Floor = Int(Value)
End Function

''
' Shifts the bits of an integer left.
'
' @param Value The value to be shifted left.
' @param Count The number of bits to shift.
' @return The newly shifted value.
'
Public Function LShift(ByVal Value As Long, ByVal Count As Long) As Long
    LShift = Helper.ShiftLeft(Value, Count)
End Function

''
' Shifts the bits of an integer right.
'
' @param Value The value to be shifted right.
' @param Count The number of bits to shift.
' @return The newly shifted value.
'
Public Function RShift(ByVal Value As Long, ByVal Count As Long) As Long
    RShift = Helper.ShiftRight(Value, Count)
End Function

''
' Divides two integers, placing the remainder in a supplied variable.
'
' @param a The dividend.
' @param b The divosor.
' @param Remainder The variable to place the remainder of the division.
' @return The quotient of the division.
Public Function DivRem(ByVal a As Long, ByVal b As Long, ByRef Remainder As Long) As Long
    DivRem = a \ b
    Remainder = a - (b * DivRem)    ' this is about 2x faster than Mod.
End Function

''
' Represents the natural logarithmic base, specified by the constant, <b><i>e</i></b>.
'
' @return <b><i>e</i></b> (Approx 2.71828182845905)
'
Public Property Get e() As Double
    e = 2.71828182845905
End Property

''
' Represents the ratio of the circumference of a circle to its diameter.
'
' @return Pi (approx 3.14159265358979)
'
Public Property Get PI() As Double
    PI = corPI
End Property

''
' Returns the maximum of two values.
'
' @param X Value to test as potential max value.
' @param Y Value to test as potential max value.
' @return The max value.
' @remarks The values must be of the same type for the comparison to be performed.
' If <b>X</b> is an <b>IComparable</b>, then the object is used to test for
' the maximum value.
' <p>If both values are equavalent then <b>X</b> is returned.</p>
'
Public Function Max(ByRef x As Variant, ByRef y As Variant) As Variant
    If IsGtOrEq(x, y) Then
        VariantCopyInd Max, x
    Else
        VariantCopyInd Max, y
    End If
End Function

''
' Returns the maximum value from a list of values.
'
' @param List The list of values to continue searching for the maximum.
' @return The maximum value found.
' @remarks The list can be an <b>Array</b>, <b>Collection</b>, or <b>ICollection</b> object.
' <p>The values must be of the same type for the comparison to be performed.
' If <b>X</b> is an <b>IComparable</b>, then the object is used to test for
' the maximum value.
'
Public Function MaxIn(ByRef List As Variant) As Variant
    Dim StartComparing  As Boolean
    Dim Value           As Variant
    
    For Each Value In List
        If StartComparing Then
            If IsGtOrEq(Value, MaxIn) Then
                Helper.MoveVariant MaxIn, Value
            End If
        Else
            Helper.MoveVariant MaxIn, Value
            StartComparing = True
        End If
    Next
End Function

''
' Returns the minimum of two values.
'
' @param x Value to test as potential min value.
' @param y Value to test as potential min value.
' @return The min value.
' @remarks The values must be of the same type for the comparison to be performed.
' If <b>X</b> is an <b>IComparable</b>, then the object is used to test for
' the minimum value.
' <p>If both values are equavalent then <b>X</b> is returned.</p>
'
Public Function Min(ByRef x As Variant, ByRef y As Variant) As Variant
    If IsLtOrEq(x, y) Then
        VariantCopyInd Min, x
    Else
        VariantCopyInd Min, y
    End If
End Function

''
' Returns the minimum value in a list of values.
'
' @param List A list of values to find the mininum value in.
' @return The mininum value found in the list.
' @remarks The list can be an <b>Array</b>, <b>Collection</b>, or <b>ICollection</b> object.
' <p>The values must be of the same type for the comparison to be performed.
' If <b>X</b> is an <b>IComparable</b>, then the object is used to test for
' the minimum value.
'
Public Function MinIn(ByRef List As Variant) As Variant
    Dim StartComparing  As Boolean
    Dim Value           As Variant
    
    For Each Value In List
        If StartComparing Then
            If IsLtOrEq(Value, MinIn) Then
                Helper.MoveVariant MinIn, Value
            End If
        Else
            Helper.MoveVariant MinIn, Value
            StartComparing = True
        End If
    Next
End Function

''
' Returns the angle from a Cosine.
'
' @param d A number representing a Cosine.
' @return The angle in radians.
'
Public Function ACos(ByVal d As Double) As Double
    ACos = Atn(-d / Sqr(-d * d + 1)) + 2 * corAtn1
End Function

''
' Returns the angle from a Sine.
'
' @param d A number representing a Sine.
' @return The angle in radians.
'
Public Function ASin(ByVal d As Double) As Double
    ASin = Atn(d / Sqr(-d * d + 1))
End Function

''
' Returns the hyperbolic cosine of the specified angle.
'
' @param d An angle in radians.
' @return The hyberbolic cosine.
'
Public Function Cosh(ByVal d As Double) As Double
    Cosh = (Exp(d) + Exp(-d)) / 2
End Function

''
' Returns the hyberbolic sine of the specified angle.
'
' @param d An angle in radians.
' @return The hyberbolic sine.
'
Public Function Sinh(ByVal d As Double) As Double
    Sinh = (Exp(d) - Exp(-d)) / 2
End Function

''
' Returns the hyberbolic tangent of the specified angle.
'
' @param d An angle in radians.
' @return The hyperbolic tangent.
'
Public Function Tanh(ByVal d As Double) As Double
    Dim PosExp As Double
    Dim NegExp As Double
    
    PosExp = Exp(d)
    NegExp = Exp(-d)
    
    Tanh = (PosExp - NegExp) / (PosExp + NegExp)
End Function

''
' Returns a value in the specified base.
'
' @param d The number to calculate the log of.
' @param NewBase The base of the log value.
' @return The log value.
'
Public Function LogBase(ByVal d As Double, ByVal NewBase As Double) As Double
    LogBase = Log(d) / Log(NewBase)
End Function

''
' Returns a log value in a base 10 log.
'
' @param d The number to calculate the log of.
' @return A base 10 log value.
'
Public Function Log10(ByVal d As Double) As Double
    Log10 = Log(d) / corLog10
End Function

''
' Returns the remainder resulting from the division of a specified number by another specified number.
'
' @param x A dividend.
' @param y A devisor.
' @return The remainder of the division.
'
Public Function IEEERemainder(ByVal x As Double, ByVal y As Double) As Double
    Dim Ret As Double
    Ret = (x - (y * Round(x / y)))
    
    If (Ret = 0#) And (x < 0#) Then
        Ret = mNegativeZero
    End If
    
    IEEERemainder = Ret
End Function

''
' Converts radians to degrees.
'
' @param Radians An angle in radians.
' @return An angle in degrees.
'
Public Function CDeg(ByVal Radians As Double) As Double
    CDeg = 180# * Radians / corPI
End Function

''
' Converts degrees to radians.
'
' @param Degrees An angle in degrees.
' @return An angle in radians.
'
Public Function CRad(ByVal Degrees As Double) As Double
    CRad = corPI * Degrees / 180#
End Function

''
' Produces the full product of two 32-bit numbers.
'
' @param a The first number to multiply.
' @param b The second number to multiply.
' @return The number containing the product of the specified numbers.
Public Function BigMul(ByVal a As Long, ByVal b As Long) As Int64
    Dim a64 As Int64
    Dim b64 As Int64
    
    If a < 0 Then
        If b < 0 Then
            a64.LowPart = -a
            b64.LowPart = -b
            Statics.Int64.MultiplyPositives a64, b64, BigMul
        Else
            a64.LowPart = a
            a64.HighPart = -1
            b64.LowPart = b
            Statics.Int64.MultiplyByNegative b64, a64, BigMul
        End If
    ElseIf b < 0 Then
        a64.LowPart = a
        b64.LowPart = b
        b64.HighPart = -1
        Statics.Int64.MultiplyByNegative a64, b64, BigMul
    Else
        a64.LowPart = a
        b64.LowPart = b
        Statics.Int64.MultiplyPositives a64, b64, BigMul
    End If
End Function

''
' Returns value clamped to the inclusive range of <i>Min</i> and <i>Max</i>.
'
' @param Value The value to be clamped.
' @param Min The lower bound of the result.
' @param Max The upper bound of the result.
' @return Returns <i>Value</i> if between <i>Min</i> and <i>Max</i>.<br/>
' -or-<br/>
' <i>Min</i> if <i>Value</i> is less than.<br/>
' -or-<br/>
' <i>Max</i> if <i>Value</i> is greater than.
Public Function Clamp(ByRef Value As Variant, ByRef Min As Variant, ByRef Max As Variant) As Variant
    Dim Value64 As Int64
    Dim Min64 As Int64
    Dim Max64 As Int64
    
    If IsAnyInt64(Value, Min, Max) Then
        Value64 = CInt64(Value)
        Min64 = CInt64(Min)
        Max64 = CInt64(Max)
        
        If Statics.Int64.Compare(Value64, Min64) < 0 Then
            Clamp = Min64
        ElseIf Statics.Int64.Compare(Value64, Max64) > 0 Then
            Clamp = Max64
        Else
            Clamp = Value
        End If
    Else
        If Value < Min Then
            Clamp = Min
        ElseIf Value > Max Then
            Clamp = Max
        Else
            Clamp = Value
        End If
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mNegativeZero = AsDouble(corNegZero)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsAnyInt64(ByRef a As Variant, ByRef b As Variant, ByRef c As Variant) As Boolean
    If IsInt64(a) Then
        IsAnyInt64 = True
    ElseIf IsInt64(b) Then
        IsAnyInt64 = True
    Else
        IsAnyInt64 = IsInt64(c)
    End If
End Function

Private Function IsGtOrEq(ByRef x As Variant, ByRef y As Variant) As Boolean
    Select Case VarType(x)
        Case vbObject
            If Not IsObject(y) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, TypeName(x)), "Y")
            
            IsGtOrEq = Comparer.Default.Compare(x, y) >= 0
        Case vbUserDefinedType
            If (Not IsInt64(x)) Or (Not IsInt64(y)) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, "Int64"))
            
            IsGtOrEq = CallFunc_T_T_Long(AddressOf Int64Compare, DataPtr(x), DataPtr(y)) >= 0
        Case VarType(y)
            IsGtOrEq = x >= y
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, TypeName(x)), "Y")
    End Select
End Function

Private Function IsLtOrEq(ByRef x As Variant, ByRef y As Variant) As Boolean
    Select Case VarType(x)
        Case vbObject
            If Not IsObject(y) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, TypeName(x)), "Y")
            
            IsLtOrEq = Comparer.Default.Compare(x, y) <= 0
        Case vbUserDefinedType
            If (Not IsInt64(x)) Or (Not IsInt64(y)) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, "Int64"))
            
            IsLtOrEq = CallFunc_T_T_Long(AddressOf Int64Compare, DataPtr(x), DataPtr(y)) <= 0
        Case VarType(y)
            IsLtOrEq = x <= y
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_MustBeOfType, TypeName(x)), "Y")
    End Select
End Function

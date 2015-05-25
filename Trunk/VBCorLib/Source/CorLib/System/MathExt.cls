VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MathExt"
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
' Module: MathExt
'

''
' Provides extended math functionality.
'
' @remarks The functions in this class are global. To access them simply call
' the functions directly as any normal function. The MathExt class does not need
' to be referenced.
'
' <pre>
' Debug.Print PI
' </pre>
'
Option Explicit

Private Const ATN_1     As Double = 0.785398163397448
Private Const LOG_10    As Double = 2.30258509299405
Private Const NEG_ZERO  As Currency = -922337203685477.5807@ - 0.0001@   ' VB won't accept -922337203685477.5808@ as a constant
Private Const C_PI      As Double = 3.14159265358979

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
'
Public Function DivRem(ByVal a As Long, ByVal b As Long, ByRef remainder As Long) As Long
    DivRem = a \ b
    remainder = a - (b * DivRem)    ' this is about 2x faster than Mod.
End Function

''
' Represents the natural logarithmic base, specified by the constant, <b><i>e</i></b>.
'
' @return <b><i>e</i></b> (Approx 2.71828182845905)
'
Public Property Get E() As Double
    E = 2.71828182845905
End Property

''
' Represents the ratio of the circumference of a circle to its diameter.
'
' @return Pi (approx 3.14159265358979)
'
Public Property Get PI() As Double
    PI = C_PI
End Property

''
' Returns the maximum of two values.
'
' @param x Value to test as potential max value.
' @param y Value to test as potential max value.
' @return The max value.
' @remarks No type checking is performed to determine the maximum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the maximum value.
' <p>If both values are equavalent then <b>x</b> is returned.</p>
'
Public Function Max(ByRef X As Variant, ByRef Y As Variant) As Variant
    On Error GoTo errTrap
    GetMax X, Y, Max
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Values could not be compared.")
End Function

''
' Returns the maximum value from a list of values.
'
' @param List The list of values to continue searching for the maximum.
' @return The maximum value found.
' @remarks The list can be an <b>Array</b>, <b>Collection</b>, or <b>ICollection</b> object.
' <p>No type checking is performed to determine the maximum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the maximum value.</p>
'
Public Function MaxIn(ByRef List As Variant) As Variant
    On Error GoTo errTrap
    
    Dim IsFirstElement As Boolean
    IsFirstElement = True
    
    Dim v As Variant
    For Each v In List
        If IsFirstElement Then
            Call VariantCopy(MaxIn, v)
            IsFirstElement = False
        Else
            Call GetMax(MaxIn, v, MaxIn)
        End If
    Next v
    
    On Error GoTo 0
    If IsFirstElement Then _
        Throw Cor.NewArgumentException("List cannot be empty.", "List")
        
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Could not compare values.", "List")
End Function

''
' Returns the minimum of two values.
'
' @param x Value to test as potential min value.
' @param y Value to test as potential min value.
' @return The min value.
' @remarks No type checking is performed to determine the minimum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the minimum value.
' <p>If both values are equavalent then <b>x</b> is returned.</p>
'
Public Function Min(ByRef X As Variant, ByRef Y As Variant) As Variant
    On Error GoTo errTrap
    GetMin X, Y, Min
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Values could not be compared.")
End Function

''
' Returns the minimum value in a list of values.
'
' @param List A list of values to find the mininum value in.
' @return The mininum value found in the list.
' @remarks The list can be an <b>Array</b>, <b>Collection</b>, or <b>ICollection</b> object.
' <p>No type checking is performed to determine the minimum of two
' values given. A string compared to an integer is a legal pair of values
' to test. The only type checked is for an <b>IComparable</b> object. If
' <b>x</b> is an <b>IComparable</b>, then the object is used to test for
' the minimum value.</p>
'
Public Function MinIn(ByRef List As Variant) As Variant
    On Error GoTo errTrap
    
    Dim IsFirstElement As Boolean
    IsFirstElement = True
    
    Dim v As Variant
    For Each v In List
        If IsFirstElement Then
            VariantCopy MinIn, v
            IsFirstElement = False
        Else
            GetMin MinIn, v, MinIn
        End If
    Next v
    
    On Error GoTo 0
    If IsFirstElement Then _
        Throw Cor.NewArgumentException("List cannot be empty.", "List")
        
    Exit Function
    
errTrap:
    Throw Cor.NewArgumentException("Could not compare values.", "List")
End Function

''
' Returns the angle from a Cosine.
'
' @param d A number representing a Cosine.
' @return The angle in radians.
'
Public Function ACos(ByVal d As Double) As Double
    ACos = Atn(-d / Sqr(-d * d + 1)) + 2 * ATN_1
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
    Cosh = (exp(d) + exp(-d)) / 2
End Function

''
' Returns the hyberbolic sine of the specified angle.
'
' @param d An angle in radians.
' @return The hyberbolic sine.
'
Public Function Sinh(ByVal d As Double) As Double
    Sinh = (exp(d) - exp(-d)) / 2
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
    
    PosExp = exp(d)
    NegExp = exp(-d)
    
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
    Log10 = Log(d) / LOG_10
End Function

''
' Returns the remainder resulting from the division of a specified number by another specified number.
'
' @param x A dividend.
' @param y A devisor.
' @return The remainder of the division.
'
Public Function IEEERemainder(ByVal X As Double, ByVal Y As Double) As Double
    Dim Ret As Double
    Ret = (X - (Y * Round(X / Y)))
    
    If (Ret = 0#) And (X < 0#) Then
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
    CDeg = 180# * Radians / C_PI
End Function

''
' Converts degrees to radians.
'
' @param Degrees An angle in degrees.
' @return An angle in radians.
'
Public Function CRad(ByVal Degrees As Double) As Double
    CRad = C_PI * Degrees / 180#
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mNegativeZero = AsDouble(NEG_ZERO)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetMin(ByRef X As Variant, ByRef Y As Variant, ByRef RetVal As Variant)
    If Not IsObject(X) Then
        If X > Y Then
            RetVal = Y
        Else
            RetVal = X
        End If
    ElseIf TypeOf X Is IComparable Then
        Dim c As IComparable
        Set c = X
        If c.CompareTo(Y) > 0 Then
            Set RetVal = Y
        Else
            Set RetVal = X
        End If
    End If
End Sub

Private Sub GetMax(ByRef X As Variant, ByRef Y As Variant, ByRef RetVal As Variant)
    If Not IsObject(X) Then
        If X < Y Then
            RetVal = Y
        Else
            RetVal = X
        End If
    ElseIf TypeOf X Is IComparable Then
        Dim c As IComparable
        Set c = X
        If c.CompareTo(Y) < 0 Then
            Set RetVal = Y
        Else
            Set RetVal = X
        End If
    End If
End Sub



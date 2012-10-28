VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PublicFunctions"
Attribute VB_GlobalNameSpace = True
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
'    Module: PublicFunctions
'

''
' Provides some helper functions.
'
Option Explicit

''
' Converts a Char value greater than 32768 to the corrisponding
' negative Integer value without overflow.
'
' @param Value The Long Ascii value to be converted.
' @return Integer version of the Long value.
' @remarks The Value can only range from -32767 to 65535, otherwise
' an <b>ArgumentOutOfRangeException</b> will be thrown.
'
Public Function LongToChar(ByVal Value As Long) As Integer
    If Value < -32767 Or Value > 65535 Then _
        Throw Cor.NewArgumentOutOfRangeException("Value must be between -32767 and 65535 inclusively.", "Value", Value)
    If Value > 32767 Then Value = Value - 65536
    LongToChar = Value
End Function

''
' Converts an Integer to the corrisponding Long Ascii char.
'
' @param Value An Integer character.
' @return A positive representation of the character.
'
Public Function CharToLong(ByVal Value As Integer) As Long
    AsWord(CharToLong) = Value
End Function

''
' Provides a way to move a variant datatype to another variant without
' the knowledge if Set needs to be used for objects.
'
' @param Destination The variable to move the variant value to.
' @param Source The variable to move the variant value from.
' @remarks This function moves the contents of the source variant to the
' destination variant. It does not make a copy. This lowers duplication
' overhead when a variant value needs to be transfered to another variant.
' <p>Any value in the destination variant will be released correctly
' before the source variants is moved into it.</p>
'
Public Sub MoveVariant(ByRef Destination As Variant, ByRef Source As Variant)
    Call Helper.MoveVariant(Destination, Source)
End Sub

''
' Provides a way to copy a variant value to another variant without the
' knowledge if Set needs to be used for objects.
'
' @param Destination The variant variable to receive a copy of the value.
' @param Source The variant variable to be copied.
' @remarks This works identical to the Win32 VariantCopyInd API function.
' <p>Any value in the destination variant will be released correctly
' before the source variants is copied into it.</p>
'
Public Sub CopyVariant(ByRef Destination As Variant, ByRef Source As Variant)
    Call VariantCopyInd(Destination, Source)
End Sub

''
' Returns the controlling IUnknown interface of the object.
'
' @param Obj The object get the controlling interface of.
' @return The controlling IUnknown interface.
' @remarks
Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

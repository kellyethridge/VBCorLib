VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: ObjectStatic
'

''
' Provides default implementations of the IObject interface.
'
Option Explicit

Private mCharBuffer As CharBuffer

''
' Returns a default hashcode for the specific value.
'
' @param Value The value to generate the hashcode for.
' @return The hashcode value.
' @remarks An object should never pass itself into this
' function. Instead, the object should implement the default
' hashcode by using ObjPtr.
' <pre>
' Public Function GetHashCode As Long
'     GetHashCode = ObjPtr(CUnk(Me))
' End Function
' </pre>
'
Public Function GetHashCode(ByRef Value As Variant) As Long
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte, vbBoolean
            GetHashCode = CLng(Value)
            
        Case vbString
            GetHashCode = HashString(Value)
        
        Case vbDouble
            With AsDLong(CDbl(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbSingle
            GetHashCode = AsLong(CSng(Value))
            
        Case vbDate
            With AsDLong(CDate(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbCurrency
            With AsDLong(CCur(Value))
                GetHashCode = .LoDWord Xor .HiDWord
            End With
            
        Case vbObject, vbDataObject
            If Value Is Nothing Then
                Exit Function
            End If
            
            If TypeOf Value Is IObject Then
                Dim Obj As IObject
                Set Obj = Value
                GetHashCode = Obj.GetHashCode
            Else
                GetHashCode = ObjPtr(CUnk(Value))
            End If
            
        Case vbDecimal
            With AsDWord(Value)
                GetHashCode = .HiWord
            End With
            With MemDLong(VarPtr(Value) + 8)
                GetHashCode = GetHashCode Xor .LoDWord Xor .HiDWord
            End With
    End Select
End Function

''
' Checks the equality of the two values.
'
' @param ObjA The first value to compare with.
' @param ObjB The second value to compare.
' @return Returns True if the value and the object are the same instance, False otherwise.
'
Public Function Equals(ByRef ObjA As Variant, ByRef ObjB As Variant) As Boolean
    Dim Obj As IObject
    
    Select Case VarType(ObjA)
        Case vbString
            If VarType(ObjB) = vbString Then
                Equals = CorString.EqualsV(ObjA, ObjB)
            End If
        Case vbDate
            Equals = Statics.cDateTime.Equals(ObjA, ObjB)
        Case vbObject
            If ObjA Is Nothing Then
                If IsObject(ObjB) Then
                    Equals = (ObjB Is Nothing)
                End If
            ElseIf TypeOf ObjA Is IObject Then
                Set Obj = ObjA
                Equals = Obj.Equals(ObjB)
            ElseIf IsObject(ObjB) Then
                If ObjB Is Nothing Then
                    Exit Function
                End If
                
                If TypeOf ObjB Is IObject Then
                    Set Obj = ObjB
                    Equals = Obj.Equals(ObjA)
                Else
                    Equals = (ObjA Is ObjB)
                End If
            End If
        Case vbNull
            Equals = IsNull(ObjB)
        Case vbEmpty
            Equals = IsEmpty(ObjB)
        Case VarType(ObjB)
            Equals = (ObjA = ObjB)
    End Select
End Function

''
' Returns a default string representation of an object.
'
' @param YourObject The object to derive the name from.
' @param YourApp the App object from the calling function.
' @return A string representation of the object.
'
Public Function ToString(ByRef Value As Variant, Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim ValueType As Long
    ValueType = VarType(Value)
    
    If ValueType = vbVariant Then
        ValueType = MemLong(MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)) And &HFF
    End If
    
    Select Case ValueType
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Dim NumberFormatter As NumberFormatInfo
            If Not Provider Is Nothing Then
                Set NumberFormatter = Provider.GetFormat(GlobalizationConstants.NumberFormatInfoTypeName)
            End If
            
            If NumberFormatter Is Nothing Then
                Set NumberFormatter = NumberFormatInfo.CurrentInfo
            End If
            
            ToString = NumberFormatter.Format(Value, Format)
            
        Case vbDate
            Dim DateFormatter As DateTimeFormatInfo
            If Not Provider Is Nothing Then
                Set DateFormatter = Provider.GetFormat(GlobalizationConstants.DateTimeFormatInfoTypeName)
            End If
            
            If DateFormatter Is Nothing Then
                Set DateFormatter = DateTimeFormatInfo.CurrentInfo
            End If
            
            ToString = DateFormatter.Format(Value, Format)
            
        Case vbObject
            If Value Is Nothing Then
                ToString = ""
            ElseIf TypeOf Value Is IFormattable Then
                Dim Formattable As IFormattable
                Set Formattable = Value
                ToString = Formattable.ToString(Format, Provider)
            ElseIf TypeOf Value Is IObject Then
                Dim Obj As IObject
                Set Obj = Value
                ToString = Obj.ToString
            Else
                ToString = TypeName(Value)
            End If
        Case vbEmpty
            ToString = "Empty"
        Case vbNull
            ToString = "Null"
        Case vbMissing
            Exit Function
        Case Else
            ToString = Value
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HashString(ByRef Value As Variant) As Long
    Dim Ptr     As Long
    Dim Length  As Long
    
    Ptr = StrPtr(Value)
    Length = MemLong(Ptr - 4) \ 2
    mCharBuffer.Buffer.pvData = Ptr
    mCharBuffer.Buffer.cElements = Length
    
    Dim i As Long
    For i = 0 To Length - 1
        HashString = ((HashString * 16) + mCharBuffer.Chars(i)) And &H3FFFFFF
    Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitChars mCharBuffer
End Sub

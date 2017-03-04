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
' Module: PublicFunctions
'

''
' Provides some helper functions.
'
Option Explicit

Private mTextOut As New StringBuilder

Public Function t(ByRef Text As String) As String
    Const vbBellChar        As Integer = 7
    Const vbBackChar        As Integer = 8
    Const vbTabChar         As Integer = 9
    Const vbCrChar          As Integer = 13
    Const vbLfChar          As Integer = 10
    Const vbVerticalTabChar As Integer = 11
    Const vbFormFeedChar    As Integer = 12
    Const vbLowerNChar      As Integer = 110
    Const vbLowerRChar      As Integer = 114
    Const vbLowerAChar      As Integer = 97
    Const vbLowerBChar      As Integer = 98
    Const vbLowerFChar      As Integer = 102
    Const vbLowerUChar      As Integer = 117
    Const vbLowerVChar      As Integer = 118
    Const vbZeroChar        As Integer = &H30
    Const vbLowerXChar      As Integer = &H78
    
    Dim Chars() As Integer
    
    mTextOut.Length = 0
    mTextOut.Capacity = LenB(Text)
    Chars = AllocChars(Text)
    
    Dim i       As Long
    Dim Length  As Long
    Dim Escaped As Boolean
    Dim Ch      As Integer
    
    Length = Len(Text)
    
    On Error GoTo Catch
    Do While i < Length
        Ch = Chars(i)
        
        If Escaped Then
            Select Case Ch
                Case corLowerTChar
                    Ch = vbTabChar
                Case vbLowerNChar
                    Ch = vbLfChar
                Case vbLowerRChar
                    Ch = vbCrChar
                Case vbLowerAChar
                    Ch = vbBellChar
                Case vbLowerBChar
                    Ch = vbBackChar
                Case vbLowerFChar
                    Ch = vbFormFeedChar
                Case vbLowerVChar
                    Ch = vbVerticalTabChar
                Case vbZeroChar
                    Ch = 0
                Case vbLowerXChar, vbLowerUChar
                    i = ReadHexChar(Chars, i + 1)
                    Escaped = False
                    GoTo Continue
            End Select
            
            mTextOut.AppendChar Ch
            Escaped = False
        ElseIf Ch = corBackslashChar Then
            Escaped = True
        Else
            mTextOut.AppendChar Ch
        End If
        
        i = i + 1
Continue:
    Loop
    
    If Escaped Then _
        Error.Format Format_UnrecognizedEscapeSequence
    
    t = mTextOut.ToString
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chars
    Return
End Function

Private Function ReadHexChar(ByRef Chars() As Integer, ByVal Index As Long) As Long
    Dim Max     As Long
    Dim Result  As Long
    Dim Ch      As Long
    Dim Count   As Long
    
    Max = UBound(Chars)
    
    If Index > Max Then _
        Error.Format Format_UnrecognizedEscapeSequence
    If Not TryHexToChar(Chars(Index), Result) Then _
        Error.Format Format_UnrecognizedEscapeSequence
    
    Index = Index + 1
    Do While Index <= Max And Count < 3
        If Not TryHexToChar(Chars(Index), Ch) Then
            Exit Do
        End If
        
        Result = Result * &H10 + Ch
        Index = Index + 1
        Count = Count + 1
    Loop
    
    mTextOut.AppendChar AsWord(Result)
    ReadHexChar = Index
End Function

Private Function TryHexToChar(ByVal Ch As Integer, ByRef Result As Long) As Boolean
    Select Case Ch
        Case 48 To 57
            Result = Ch - 48
        Case 97 To 102
            Result = Ch - 87
        Case 65 To 70
            Result = Ch - 55
        Case Else
            Exit Function
    End Select
    
    TryHexToChar = True
End Function

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
        Throw Cor.NewArgumentOutOfRangeException("Value", "Value must be between -32767 and 65535 inclusively.")
    If Value > 32767 Then Value = Value - 65536
    LongToChar = Value
End Function

''
' Returns character value from a string at the specified index.
'
' @param s The string to return the character from.
' @param Index The position of the character in the string.
' @return An integer representing the character.
' @remarks The index is zero based.
'
Public Function CharAt(ByRef s As String, ByVal Index As Long) As Integer
    If Index < 0 Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_NeedNonNegNum
    If Index >= Len(s) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    CharAt = MemWord(StrPtr(s) + Index * 2)
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
    Helper.MoveVariant Destination, Source
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
    VariantCopyInd Destination, Source
End Sub

''
' Returns the controlling IUnknown interface of the object.
'
' @param Obj The object get the controlling interface of.
' @return The controlling IUnknown interface.
'
Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

''
' Returns the instance as type Object.
'
' @param Obj The instance to convert.
' @return Returns the instance as type Object.
'
Public Function CObj(ByVal Obj As Object) As Object
    Set CObj = Obj
End Function


Public Function GetCollectionVersion(ByVal Obj As Object) As Long
    GetCollectionVersion = Enumeration.GetCollectionVersion(Obj)
End Function

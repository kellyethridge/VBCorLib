VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: StringBuilder
'

''
' Allows easy manipulation, concatenation and removal of strings and characters.
'
' @remarks
' This class can be used to quickly minpulate strings. It allows for easy concatenation,
' removal and modification of the current underlying string value.
'
' @see Constructors
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Option Explicit
Implements IObject

Private Const DefaultFillChar   As Integer = vbSpace
Private Const DefaultCapacity   As Long = 512
Private Const OpenBraceChar     As Long = 123
Private Const CloseBraceChar    As Long = 125
Private Const BytesPerChar      As Long = 2


' module wide scope
Private mStringPtr  As Long
Private mStringSA   As SafeArray1d
Private mString()   As Integer
Private mCapacity   As Long
Private mLength     As Long
Private mFillChar   As Integer

' used by Replace
Private OldStr()    As Integer
Private BufStr()    As Integer
Private OldStrSA    As SafeArray1d
Private BufStrSA    As SafeArray1d

' used by AppendFormat
Private mChars      As WordBuffer


''
' Appends a string to the current string value.
'
' @param Value Value to be appended.
' @param StartIndex A zero-based index of the first character to start appending from in the string.
' @param Count The number of characters from the string to append.
' @return This instance of StringBuilder.
' @remarks This version of Append attempts to convert to the value using the
' <b>Convert.ToString</b> method. This will allow for objects that impelement
' <b>IObject</b> or <b>IFormattable</b> to be converted to a string value.
'
Public Function Append(ByRef Value As Variant, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    Dim ElemIndex As Long
    Dim CharCount As Long
    Dim s As String
    
    s = Convert.ToString(Value)
    ElemIndex = GetOptionalLong(StartIndex, 0)
    CharCount = GetOptionalLong(Count, Len(s) - ElemIndex)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index), "StartIndex")
    If CharCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count")
    If ElemIndex + CharCount > Len(s) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
        
    If mLength + CharCount > mCapacity Then
        EnsureCapacity mLength + CharCount
    End If
    
    CopyMemory ByVal mStringPtr + mLength * 2, ByVal StrPtr(s) + ElemIndex * BytesPerChar, CharCount * BytesPerChar
    mLength = mLength + CharCount
    
    Set Append = Me
End Function

''
' Appends a character a specified number of times to the end of string value.
'
' @param char The AscW character code to append.
' @param count The number of times to append the character.
' @return This instance of StringBuilder.
' @remarks The character is a unicode value from -32767 to 65535. For values that
' are greater than 32767 (&H7FFF), 65536 (&H10000) is subtracted from it to map the
' character to a 16-bit Integer without overflowing.
'
Public Function AppendChar(ByVal Char As Long, Optional ByVal Count As Long = 1) As StringBuilder
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count cannot be negative.", "Count", Count)
    If Char < -&H7FFF& Or Char > &HFFFF& Then _
        Throw Cor.NewArgumentOutOfRangeException("Char must be between -32767 and 65535 inclusively.", "Char", Char)
    
    If Char > &H7FFF& Then Char = Char - &H10000
    
    If mLength + Count > mCapacity Then
        EnsureCapacity mLength + Count
    End If
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If Count = 1 Then
        mString(mLength) = Char
    Else
        Fill mLength, Char, Count
    End If
    mLength = mLength + Count
    Set AppendChar = Me
End Function

''
' Appends an array of characters to the end of the string value.
'
' @param Chars The array of characters to be appended.
' @param Startindex The index from Chars to start appending.
' @param CharCount The number of characters to append.
' @return This instance of StringBuilder.
'
Public Function AppendChars(ByRef Chars() As Integer, Optional ByRef StartIndex As Variant, Optional ByRef CharCount As Variant) As StringBuilder
    If SAPtr(Chars) = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Chars")
    If SafeArrayGetDim(SAPtr(Chars)) > 1 Then _
        Throw Cor.NewRankException(Resources.GetString(Rank_MultiDimNotSupported))
        
    Dim LocalCharCount As Long
    Dim ElemIndex As Long
    ElemIndex = GetOptionalLong(StartIndex, LBound(Chars))
    LocalCharCount = GetOptionalLong(CharCount, UBound(Chars) - ElemIndex + 1)
    
    If ElemIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "StartIndex")
    If ElemIndex + LocalCharCount - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
    
    If LocalCharCount > 0 Then
        If mLength + LocalCharCount > mCapacity Then
            EnsureCapacity mLength + LocalCharCount
        End If
        
        CopyMemory mString(mLength), Chars(ElemIndex), LocalCharCount * BytesPerChar
        mLength = mLength + LocalCharCount
    End If
    
    Set AppendChars = Me
End Function

''
' Appends a string with formatted arguments.
'
' @param format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\..\Includes\StringBuilder.AppendFormat.txt"
Public Function AppendFormat(ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16)
    InternalAppendFormat Nothing, Format, vArgs
    Set AppendFormat = Me
End Function

''
' Appends a string with formatted arguments using the supplied provider.
'
' @param provider A custom formatting object that formats the arguments.
' @param Format The string to append that contains formatting information.
' @param args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks
' <p>The format string can embed indexes into the args parameter to indicate where
' a value should be placed. The way to indicate an index is through an argument index enclosed
' in braces {N}. N is the parameter index beginning with zero.</p>
' <br><p>Some arguments may handle additional formatting commands. Please refer to their
' documentation for more information.</p>
' <br><p>The supplied provider can format specific arguments unknown to the system. It can also
' delegate formatting of system-known arguments to another StringBuilder (usually cString.Format)
' to provide default formatting of known types and format only specific types.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function AppendFormatEx(ByRef Provider As IFormatProvider, ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20)
    InternalAppendFormat Provider, Format, vArgs
    Set AppendFormatEx = Me
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of StringBuilder.
'
Public Function AppendLine(Optional ByRef s As String) As StringBuilder
    AppendQuick s
    AppendQuick vbCrLf
    Set AppendLine = Me
End Function

''
' This function is designed to append strings quickly.
'
' @param s The string value to be appended.
' @remarks This method appends the entire string to the end of the
' StringBuilder. There is no return reference to the StringBuilder.
'
Public Sub AppendQuick(ByRef s As String)
    Dim CharCount As Long
    CharCount = Len(s)
    
    If CharCount > 0 Then
        If mLength + CharCount > mCapacity Then
            EnsureCapacity mLength + CharCount
        End If
        
        CopyMemory ByVal mStringPtr + mLength * 2, ByVal StrPtr(s), CharCount * BytesPerChar
        mLength = mLength + CharCount
    End If
End Sub

''
' Copies a set of characters from the StringBuilder into a character array.
'
' @param sourceindex The starting index in the StringBuilder to be copying from (0-based).
' @param Destination The character array to copy to.
' @param destinationindex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
'
Public Sub CopyTo(ByVal SourceIndex As Long, ByRef Destination() As Integer, ByVal DestinationIndex As Long, ByVal Count As Long)
    If SourceIndex < 0 Or SourceIndex > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index), "SourceIndex")
    If SAPtr(Destination) = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Destination")
    If DestinationIndex < LBound(Destination) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "DestinationIndex")
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count")
    If SourceIndex + Count > mLength Then _
        Throw Cor.NewArgumentException(Resources.GetString(Arg_LongerThanSrcString))
    If DestinationIndex + Count - 1 > UBound(Destination) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_OffsetOut))
    
    If Count > 0 Then
        CopyMemory Destination(DestinationIndex), ByVal mStringPtr + SourceIndex * BytesPerChar, Count * BytesPerChar
    End If
End Sub

''
' Returns the buffer size of the string value.
'
' @return Value indicating the amount of memory allocated in characters.
'
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the amount of memory allocated in characters.
'
' @param RHS The number of characters to allocate.
'
Public Property Let Capacity(ByVal RHS As Long)
    If RHS < mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Cannot set capacity less than the length of the current instance.", "Capacity", RHS)
    
    If RHS <> mCapacity Then
        Dim Ptr As Long
        Ptr = CoTaskMemAlloc(RHS * BytesPerChar)
        If Ptr = vbNullPtr Then _
            Throw New OutOfMemoryException
            
        CopyMemory ByVal Ptr, ByVal mStringPtr, mLength * BytesPerChar
        CoTaskMemFree mStringPtr
        mStringPtr = Ptr
        mStringSA.pvData = Ptr
        mCapacity = RHS
    End If
End Property

''
' Returns a character from the string value.
'
' @param Index The index from the string value to return the character.
' @return The character in the string at the specified index.
' @remarks Index is zero-base.
'
Public Property Get Chars(ByVal Index As Long) As Integer
Attribute Chars.VB_UserMemId = 0
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Chars = mString(Index)
End Property

''
' Sets a character in the string value.
'
' @param index The index into the string value to set the character.
' @param char The character to set.
' @remarks index is zero-based.
'
Public Property Let Chars(ByVal Index As Long, ByVal Char As Integer)
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "Index", Index)
    
    mString(Index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param requiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
'
Public Function EnsureCapacity(ByVal RequiredCapacity As Long) As Long
    If RequiredCapacity > mCapacity Then
        Dim NewCapacity As Long
        NewCapacity = mCapacity * 2
        If NewCapacity >= RequiredCapacity Then
            Capacity = NewCapacity
        Else
            Capacity = RequiredCapacity
        End If
    End If
    
    EnsureCapacity = mCapacity
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = Object.Equals(Me, Value)
End Function

''
' Returns the fill character used to pad empty space.
'
' @return An Integer representing the AscW of the character.
'
Public Property Get FillCharacter() As Variant
    FillCharacter = mFillChar
End Property

''
' Sets the fill character used to pad empty space.
'
' @param RHS The character as either a String or Integer value.
' @remarks The new value can be either String, in which case the
' first characters is used, or it can be a number.
'
Public Property Let FillCharacter(ByVal RHS As Variant)
    Select Case VarType(RHS)
        Case vbLong, vbInteger, vbByte
            mFillChar = AscW(ChrW$(RHS))
        Case vbString
            mFillChar = AscW(RHS)
        Case Else
            Throw Cor.NewInvalidCastException("Fill Character can be either a string or number.")
    End Select
End Property

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Inserts a string into the current string value.
'
' @param index The index in the string value to start inserting the new string.
' @param value The string to be inserted.
' @param count The number of times to insert the string.
' @return This instance of StringBuilder
'
Public Function Insert(ByVal Index As Long, ByRef Value As Variant, Optional ByVal Count As Long = 1) As StringBuilder
    If Index < 0 Or Index > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index), "Index", Index)
    If Count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count", Count)
    
    Dim s As String
    s = Convert.ToString(Value)
    Dim l As Long
    l = Len(s)
    If l > 0 And Count > 0 Then
        InsertSpace Index, l * Count
        l = l * BytesPerChar
        Index = Index * BytesPerChar
        Do While Count > 0
            CopyMemory ByVal mStringPtr + Index, ByVal StrPtr(s), l
            Index = Index + l
            Count = Count - 1
        Loop
    End If
    Set Insert = Me
End Function

''
' Inserts an array of characters into the string value.
'
' @param Index The index in the string value to start inserting.
' @param Chars The characters to be inserted.
' @param Startindex The index in the Characters to start inserting from.
' @param CharCount The number of characters to be inserted.
' @return This instance of StringBuilder.
'
Public Function InsertChars(ByVal Index As Long, ByRef Chars() As Integer, Optional ByRef StartIndex As Variant, Optional ByRef CharCount As Variant) As StringBuilder
    If Index < 0 Or Index > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index), "Index")
    If SAPtr(Chars) = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Resources.GetString(ArgumentNull_Array), "Chars")
    
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = GetOptionalLong(StartIndex, LBound(Chars))
    ElemCount = GetOptionalLong(CharCount, UBound(Chars) - ElemIndex + 1)
    
    If ElemIndex < LBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_LBound), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "CharCount")
    If ElemIndex + ElemCount - 1 > UBound(Chars) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index), "StartIndex")
        
    If ElemCount > 0 Then
        InsertSpace Index, ElemCount
        CopyMemory mString(Index), Chars(ElemIndex), ElemCount * BytesPerChar
    End If
    Set InsertChars = Me
End Function

''
' Returns the length of the current string value.
'
' @return The length of the string value.
'
Public Property Get Length() As Long
    Length = mLength
End Property

''
' Sets the length of the current string value.
'
' @param RHS The length of the string value.
' @remarks If the new value is longer than the current length, then the
' new space is padded with FillCharacter (default 32).
'
Public Property Let Length(ByVal RHS As Long)
    If RHS < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NegativeLength), "Length", Length)
    
    If RHS > mLength Then
        If RHS > mCapacity Then
            EnsureCapacity RHS
        End If
        Fill mLength, mFillChar, mCapacity - mLength
    End If
    mLength = RHS
End Property

''
' Removes a number of characters from the string value.
'
' @param startindex The index in the string value to start removing from.
' @param count The number of characters to remove from the string value.
' @return This instance of StringBuilder.
' @remarks startindex is zero-based.
'
Public Function Remove(ByVal StartIndex As Long, ByVal Length As Long) As StringBuilder
    If StartIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_StartIndex), "StartIndex")
    If Length < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NegativeLength), "Length")
    If StartIndex + Length > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
        
    If Length > 0 Then
        CopyMemory ByVal mStringPtr + StartIndex * 2, ByVal mStringPtr + (StartIndex + Length) * BytesPerChar, (mLength - (StartIndex + Length)) * BytesPerChar
        mLength = mLength - Length
    End If
    Set Remove = Me
End Function

''
' Replaces a substring in the string value with a new string.
'
' @param OldValue The substring to be replaced.
' @param NewValue The string to replace the old string with.
' @param StartIndex The index of the start of the substring in the StringBuilder object.
' @param count The number of characters in the StringBuilder object substring.
' @return This instance of StringBuilder.
' @remarks Use index and count to replace the old value with a substring of the new value. index is zero-based.
'
Public Function Replace(ByRef OldValue As String, ByRef NewValue As String, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant) As StringBuilder
    If Len(OldValue) = 0 Then _
        Throw Cor.NewArgumentException(Resources.GetString(Argument_EmptyName), "OldValue")
    
    Dim ElemStart As Long
    Dim ElemCount As Long
    ElemStart = GetOptionalLong(StartIndex, 0)
    ElemCount = GetOptionalLong(Count, mLength - ElemStart)
    
    If ElemStart < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_StartIndex), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Count")
    If ElemStart + ElemCount > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
    
    Dim MatchCount      As Long
    Dim MatchIndexes()  As Long
    MatchCount = FindMatches(OldValue, ElemStart, ElemCount, MatchIndexes)

    ' If we have matches then we need to replace them.
    If MatchCount > 0 Then
        Dim NewLength As Long
        NewLength = mLength - ((Len(OldValue) - Len(NewValue)) * MatchCount)

        Select Case NewLength
            Case mLength:       ReplaceEqualLength NewValue, MatchIndexes, MatchCount
            Case Is < mLength:  ReplaceSmallerLength NewValue, MatchIndexes, MatchCount, Len(OldValue)
            Case Else:          ReplaceLargerLength NewValue, MatchIndexes, MatchCount, Len(OldValue), NewLength
        End Select

        mLength = NewLength
    End If

    Set Replace = Me
End Function

''
' Returns the current version of the string value.
'
' @param startindex The index to start the return substring from.
' @param length The number of characters to return in the string.
' @return A string or substring representing the internal string of the builder.
' @remarks startindex is zero-based.
'
Public Function ToString(Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As String
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    ElemIndex = GetOptionalLong(StartIndex, 0)
    ElemCount = GetOptionalLong(Length, mLength)
    
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_StartIndex), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NegativeLength), "Length")
    If ElemIndex + ElemCount > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_Index))
    
    If ElemCount = 0 Then
        ToString = ""
    Else
        ToString = SysAllocStringLen(mStringPtr + ElemIndex * BytesPerChar, ElemCount)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByRef StartIndex As Variant, ByRef Length As Variant, ByVal Capacity As Long)
    Dim ElemIndex As Long
    Dim ElemCount As Long
    ElemIndex = GetOptionalLong(StartIndex, 0)
    ElemCount = GetOptionalLong(Length, Len(s) - ElemIndex)
        
    If ElemIndex < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_StartIndex), "StartIndex")
    If ElemCount < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Length")
    If ElemIndex + ElemCount > Len(s) Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_IndexLength))
    If Capacity < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Resources.GetString(ArgumentOutOfRange_NeedNonNegNum), "Capacity")
        
    If Capacity = 0 Then
        Capacity = DefaultCapacity
    End If
    
    If ElemCount > Capacity Then
        Capacity = ElemCount
    End If
    
    If Capacity <> mCapacity Then
        CoTaskMemFree mStringPtr
        mStringPtr = CoTaskMemAlloc(Capacity * BytesPerChar)
        mCapacity = Capacity
    End If
    
    
    If ElemCount > 0 Then
        CopyMemory ByVal mStringPtr, ByVal StrPtr(s) + ElemIndex * BytesPerChar, ElemCount * BytesPerChar
        mLength = ElemCount
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mFillChar = DefaultFillChar
    mStringPtr = CoTaskMemAlloc(DefaultCapacity * BytesPerChar)
    If mStringPtr = vbNullPtr Then _
        Throw New OutOfMemoryException
    mCapacity = DefaultCapacity
    
    With OldStrSA
        .cbElements = BytesPerChar
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    SAPtr(OldStr) = VarPtr(OldStrSA)
    BufStrSA = OldStrSA
    SAPtr(BufStr) = VarPtr(BufStrSA)
    mStringSA = OldStrSA
    SAPtr(mString) = VarPtr(mStringSA)
    mStringSA.pvData = mStringPtr
    
    InitWordBuffer mChars, 0, &H7FFFFFFF
End Sub

Private Sub Class_Terminate()
    CoTaskMemFree mStringPtr
    SAPtr(OldStr) = 0
    SAPtr(BufStr) = 0
    SAPtr(mString) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Fill(ByVal Index As Long, ByVal Char As Integer, ByVal Count As Long)
    Dim i As Long
    
    For i = Index To Index + Count - 1
        mString(i) = Char
    Next i
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException(Resources.GetString(Format_InvalidString))
End Sub

Private Sub InsertSpace(ByVal Index As Long, ByVal size As Long)
    Dim NewLength As Long
    NewLength = mLength + size
    
    EnsureCapacity NewLength
    
    If Index < mLength Then
        ' Create space in the string. The new space is uninitialized.
        CopyMemory ByVal mStringPtr + (Index + size) * BytesPerChar, ByVal mStringPtr + Index * BytesPerChar, (mLength - Index) * BytesPerChar
    End If
    
    mLength = NewLength
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend helper methods
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' We can supply a custom formatter through the IFormatProvider interface.
' This article lead me to believe that the StringBuilder.AppendFormat function
' can use an ICustomFormatter object, even though information about it is
' scarce in the MSDN.
'
' "http://www.codeproject.com/csharp/custstrformat.asp"
'
Friend Sub InternalAppendFormat(ByRef Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant)
    Dim CustomFormatter     As ICustomFormatter
    Dim HasCustomFormatter  As Boolean
    Dim Index               As Long
    Dim Max                 As Long
    Dim SubStringStart      As Long
    
    ' See if the user wants to format anything their own way.
    If Not Provider Is Nothing Then
        Set CustomFormatter = Provider.GetFormat("ICustomFormatter")
        HasCustomFormatter = Not (CustomFormatter Is Nothing)
    End If
    
    Max = Len(Format)
    mChars.SA.pvData = StrPtr(Format)
    mChars.SA.cElements = Max
    
    ' Parse normal substring.
    Do While Index < Max
        Select Case mChars.Data(Index)
            Case OpenBraceChar
                ' potential escape or format info.
                
                If Index > SubStringStart Then
                    ' If we have some characters, just append them now
                    ' even if the "{" may be escaped. Most times it won't be.
                    AppendChars mChars.Data, SubStringStart, Index - SubStringStart
                End If
                
                Index = Index + 1
                
                ' We test for escape by checking the next character for a "{".
                ' But if we're out of characters, then we ended with a "{",
                ' and that is an invalid format.
                If Index = Max Then FormatError
                
                ' We have enough characters to test for possible escape.
                If mChars.Data(Index) = OpenBraceChar Then
                    ' escape it by setting the start of the substring to the second one.
                    SubStringStart = Index
                Else
                    Dim ArgIndex        As Long
                    Dim ArgWidth        As Long
                    Dim ArgFormatting   As String
                    
                    ' must be format info within a "{}", so call the parser for it.
                    Index = ParseArgumentFormatting(Index, ArgIndex, ArgWidth, ArgFormatting)
                    If ArgIndex < 0 Or ArgIndex > UBound(Args) Then _
                        Throw Cor.NewFormatException("The argument index must be from 0 to " & UBound(Args) & ".")
                        
                    ' append formatted argument here.
                    Dim s As String
                    
                    ' Set this because we are in a loop and always
                    ' need to begin with a vbNullString.
                    s = vbNullString
                    
                    ' Give a custom formatter first shot at formatting.
                    If HasCustomFormatter Then
                        s = CustomFormatter.Format(ArgFormatting, Args(ArgIndex), Provider)
                    End If
                    
                    ' We distinguish between a null string and an empty string.
                    ' Null string indicates that the value was not formatted.
                    If StrPtr(s) = vbNullPtr Then
                        s = Convert.ToString(Args(ArgIndex), ArgFormatting, Provider)
                    End If
                    
                    ' Add leading padding.
                    If ArgWidth > Len(s) Then
                        AppendChar vbSpace, ArgWidth - Len(s)
                    End If
                    
                    ' Finally append the formatted value to the string.
                    AppendQuick s
                    
                    ' Add trailing padding.
                    If -ArgWidth > Len(s) Then
                        AppendChar vbSpace, (-ArgWidth) - Len(s)
                    End If
                    
                    SubStringStart = Index + 1
                    
                End If
            Case CloseBraceChar
                ' must be escaped or else it is a format error.
                
                Index = Index + 1
                
                ' Need enough characters to test for escape.
                If Index = Max Then FormatError
                
                If mChars.Data(Index) = CloseBraceChar Then
                    ' escape it and add the substring to this point.
                    ' The substring ends one character back because we
                    ' don't want to include both "}}" in the substring.
                    If Index - 1 > SubStringStart Then
                        ' Append the current substring.
                        AppendChars mChars.Data, SubStringStart, Index - SubStringStart - 1
                    End If
                    SubStringStart = Index
                Else
                    ' we failed to escape it, so there was no matching "{".
                    FormatError
                End If
        End Select
        Index = Index + 1
    Loop
    
    If Index > SubStringStart Then
        ' If we have some characters, just append them now.
        AppendChars mChars.Data, SubStringStart, Index - SubStringStart
    End If
End Sub

' We must be inside "{}". We must end with a matching "}" or it is a format error.
'
' @param Index The index of the first character after the "{".
' @param ArgIndex The index of the argument in the Args array.
' @param ArgWidth The width of the column the argument will be displayed in.
' @param ArgFormat The format information used to determine formatting for the argument.
' @return The index of the terminating "}".
'
Private Function ParseArgumentFormatting(ByVal Index As Long, ByRef ArgIndex As Long, ByRef ArgWidth As Long, ByRef ArgFormat As String) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    ' Get the argument index.
    ArgIndex = ParseNumber(Index)
    
    Index = ConsumeSpaces(Index)
    
    With mChars
        ' Check if a column size is being defined.
        If .Data(Index) = vbComma Then
            Index = ConsumeSpaces(Index + 1)
            
            ' Get the column width and aligment. Remember
            ' negative values mean align left in the column.
            ArgWidth = ParseNumber(Index)
            
            Index = ConsumeSpaces(Index)
        Else
            ' Make sure they know there was no width specified.
            ArgWidth = 0
        End If
        
        ' Check if formatting is being specified.
        If .Data(Index) = vbColon Then
            Dim StartIndex As Long
            
            ' everything from here to the "}" is part of the argument formatting.
            Index = Index + 1
            StartIndex = Index
            Do
                ' can never reach the end until after a "}".
                If Index = Max Then FormatError
                If .Data(Index) = CloseBraceChar Then
                    ArgFormat = SysAllocStringLen(VarPtr(.Data(StartIndex)), Index - StartIndex)
                    Exit Do
                End If
                Index = Index + 1
            Loop
        Else
            ' Make sure we let them know that no formatting was supplied.
            ArgFormat = vbNullString
        End If
        
        ' if we get here we should have reached the closing "}".
        If .Data(Index) <> CloseBraceChar Then FormatError
    End With
    ParseArgumentFormatting = Index
End Function

''
' Advances the index to the next non-space character.
'
' @param The index to the first space to be consumed.
' @return The index of the first character after the contiguous
' stream of spaces have been consumed.
'
Private Function ConsumeSpaces(ByVal Index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Do
        ' Can never hit the end consuming spaces.
        If Index = Max Then FormatError
        If mChars.Data(Index) <> vbSpace Then Exit Do
        Index = Index + 1
    Loop
    
    ConsumeSpaces = Index
End Function

''
' Parses out a positive or negative number.
'
' @param Index The index to the first digit or the "-" preceeding the first digit.
' @return The number that was parsed.
' @remarks The Index is passed by Reference and will be altered.
'
Private Function ParseNumber(ByRef Index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Dim Sign As Long
    Sign = 1
    
    Dim FirstCharIndex As Long
    FirstCharIndex = Index
    
    Dim Number      As Long
    Dim NumberFound As Boolean
    Do
        ' We can never hit the end parsing a number.
        If Index = Max Then FormatError
        
        Dim ch As Integer
        ch = mChars.Data(Index)
        Select Case ch
            Case vbZero To vbNine
                ' Keep adding digits to the number
                Number = Number * 10 + ch - vbZero
                NumberFound = True
            Case vbMinus
                ' The "-" can only be the first character.
                If Index <> FirstCharIndex Then FormatError
                Sign = -1
            Case Else
                ' we have reached the end of the number, so exit.
                Exit Do
        End Select
        Index = Index + 1
    Loop
    
    If Not NumberFound Then FormatError
    
    ParseNumber = Number * Sign
End Function

''
' Loops through the string buffer finding all matching occurrences of the old value
' and storing the index for each match found.
'
' Returns the number of matches found.
Private Function FindMatches(ByRef OldValue As String, ByVal StartIndex As Long, ByVal Count As Long, ByRef MatchIndexes() As Long) As Long
    OldStrSA.pvData = StrPtr(OldValue)
    OldStrSA.cElements = Len(OldValue)
    
    ' We'll cache this since we'll be hitting it constantly.
    Dim FirstOldValueChar As Long
    FirstOldValueChar = OldStr(0)
    
    ' Don't need to search at the end if the old value
    ' can't fit. If the end matches the old value, then
    ' the inner loop will finish out the matching.
    Count = Count - OldStrSA.cElements + 1
    
    Dim MatchCount As Long
    ReDim MatchIndexes(0 To mLength)
    
    ' Loop through the current StringBuilder string trying to find
    ' all the matches against the old value to be replaced. If a
    ' match is found, then the index to that match is stored for
    ' future usage when replacing the matched old values with the new.
    Do While Count > 0
        ' Found the start of a possible match.
        If mString(StartIndex) = FirstOldValueChar Then
            ' We'll be opitmistic and assume a match.
            Dim Matched As Boolean
            Matched = True
            
            ' We optimistically save the index now
            ' because we'll be changing it while
            ' matching occurs.
            MatchIndexes(MatchCount) = StartIndex
            
            ' Start with the second character.
            Dim i As Long
            For i = 1 To OldStrSA.cElements - 1
                If mString(StartIndex + i) <> OldStr(i) Then
                    Matched = False
                    Exit For
                End If
            Next i
                
            If Matched Then
                MatchCount = MatchCount + 1
                StartIndex = StartIndex + OldStrSA.cElements - 1
                Count = Count - OldStrSA.cElements + 1
            End If
        End If
        
        StartIndex = StartIndex + 1
        Count = Count - 1
    Loop

    FindMatches = MatchCount
End Function

''
' Optimized to simply copy the new value over the old value within the original buffer.
'
' @remarks This is used when the old and new values are the same length, resulting in
' no change in the size of the final string.
'
Private Sub ReplaceEqualLength(ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long)
    Dim NewValuePtr As Long
    NewValuePtr = StrPtr(NewValue)
    
    Dim AmountToCopy As Long
    AmountToCopy = LenB(NewValue)
    
    Dim i As Long
    For i = 0 To MatchCount - 1
        CopyMemory mString(MatchIndexes(i)), ByVal NewValuePtr, AmountToCopy
    Next i
End Sub

''
' This will simply use the original string buffer as the source and destination buffers.
'
' @remarks Since the resulting string will be smaller, there is no fear of overwritting
' any original text that needs to end up in the final string.
'
Private Sub ReplaceSmallerLength(ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long, ByVal OldValueLength As Long)
    ReplaceMatches mStringPtr, mStringPtr, NewValue, MatchIndexes, MatchCount, OldValueLength
End Sub

''
' A new buffer will need to be allocated because the final string will be longer than the original.
'
' @remarks This will use the original string buffer as the source and a newly
' allocated buffer as the destination. Once the replacement has been completed,
' the newly allocated buffer will become the string buffer and the string buffer
' memory will be released.
'
Private Sub ReplaceLargerLength(ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long, ByVal OldValueLength As Long, ByVal NewLength As Long)
    Dim NewCapacity As Long
    NewCapacity = MathExt.Max(NewLength, mCapacity)
    
    Dim NewBuffer As Long
    NewBuffer = CoTaskMemAlloc(NewCapacity * BytesPerChar)
    If NewBuffer = vbNullPtr Then Throw New OutOfMemoryException
    
    ReplaceMatches mStringPtr, NewBuffer, NewValue, MatchIndexes, MatchCount, OldValueLength
    
    CoTaskMemFree mStringPtr
    mStringPtr = NewBuffer
    mStringSA.pvData = NewBuffer
    mCapacity = NewCapacity
End Sub

''
' Performs the replacement of old values with new values, copying the unaltered text from the
' source buffer and placing the final unaltered text plus newly replaced text in the destination buffer.
'
Private Sub ReplaceMatches(ByVal lpSourceBuffer As Long, ByVal lpDestBuffer As Long, ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long, ByVal OldValueLength As Long)
    Dim NewValuePtr As Long
    NewValuePtr = StrPtr(NewValue)
    
    Dim NewValueLen As Long
    NewValueLen = LenB(NewValue)
    
    Dim OriginalIndex   As Long
    Dim AmountToCopy    As Long
    Dim i               As Long
    For i = 0 To MatchCount - 1
        AmountToCopy = MatchIndexes(i) - OriginalIndex
        
        If AmountToCopy > 0 Then
            CopyMemory ByVal lpDestBuffer, ByVal lpSourceBuffer + (OriginalIndex * BytesPerChar), AmountToCopy * BytesPerChar
            lpDestBuffer = lpDestBuffer + AmountToCopy * 2
        End If
        
        If NewValueLen > 0 Then
            CopyMemory ByVal lpDestBuffer, ByVal NewValuePtr, NewValueLen
            lpDestBuffer = lpDestBuffer + NewValueLen
        End If
        
        OriginalIndex = OriginalIndex + AmountToCopy + OldValueLength
    Next i
    
    If OriginalIndex < mLength Then
        CopyMemory ByVal lpDestBuffer, ByVal lpSourceBuffer + (OriginalIndex * BytesPerChar), (mLength - OriginalIndex) * BytesPerChar
    End If
End Sub

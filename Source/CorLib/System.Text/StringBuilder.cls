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

Private Const DefaultFillChar   As Integer = vbSpaceChar
Private Const DefaultCapacity   As Long = 128

Private mString()   As Integer
Private mCapacity   As Long
Private mLength     As Long
Private mFillChar   As Integer


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
    Select Case VarType(Value)
        Case vbIntegerArray
            AppendCharsValue Value, StartIndex, Count
        Case vbString
            AppendStringValue Value, StartIndex, Count
        Case Else
            AppendString Object.ToString(Value)
    End Select
    
    Set Append = Me
End Function

''
' Appends a Unicode character a specified number of times to the end of string value.
'
' @param Char The Unicode character code to append.
' @param Count The number of times to append the character.
' @return This instance of StringBuilder.
'
Public Function AppendChar(ByVal Char As Long, Optional ByVal Count As Long = 1) As StringBuilder
    Dim LocalChar As Integer
    
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    LocalChar = AsWord(Char)
    EnsureCapacity mLength + Count
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If Count = 1 Then
        mString(mLength) = LocalChar
    Else
        Fill mLength, LocalChar, Count
    End If
    
    mLength = mLength + Count
    Set AppendChar = Me
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
' @include "..\DocIncludes\StringBuilder.AppendFormat.html"
Public Function AppendFormat(ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    InternalAppendFormat Nothing, Format, SwappedArgs
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
' delegate formatting of system-known arguments to another StringBuilder (usually CorString.Format)
' to provide default formatting of known types and format only specific types.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
'
Public Function AppendFormatEx(ByRef Provider As IFormatProvider, ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(20)
    InternalAppendFormat Provider, Format, SwappedArgs
    Set AppendFormatEx = Me
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of StringBuilder.
'
Public Function AppendLine(Optional ByRef s As String) As StringBuilder
    If LenB(s) > 0 Then
        AppendString s
    End If
    
    AppendString vbCrLf
    Set AppendLine = Me
End Function

''
' This function is designed to append strings quickly.
'
' @param s The string value to be appended.
' @remarks This method appends the entire string to the end of the
' StringBuilder. There is no return reference to the StringBuilder.
'
Public Sub AppendString(ByRef s As String)
    AppendCore StrPtr(s), 0, Len(s)
End Sub

''
' Copies a set of characters from the StringBuilder into a character array.
'
' @param SourceIndex The starting index in the StringBuilder to be copying from (0-based).
' @param Destination The character array to copy to.
' @param DestinationIndex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
'
Public Sub CopyTo(ByVal SourceIndex As Long, ByRef Destination() As Integer, ByVal DestinationIndex As Long, ByVal Count As Long)
    If SourceIndex < 0 Or SourceIndex > mLength Then _
        Error.ArgumentOutOfRange "SourceIndex", ArgumentOutOfRange_Index
    If SAPtr(Destination) = vbNullPtr Then _
        Error.ArgumentNull "Destination", ArgumentNull_Array
    If DestinationIndex < LBound(Destination) Then _
        Error.ArgumentOutOfRange "DestinationIndex", ArgumentOutOfRange_LBound
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If SourceIndex + Count > mLength Then _
        Error.Argument Arg_LongerThanSrcString
    If DestinationIndex + Count - 1 > UBound(Destination) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_OffsetOut
    
    If Count > 0 Then
        CopyMemory Destination(DestinationIndex), mString(SourceIndex), Count * vbSizeOfChar
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
' @param Value The number of characters to allocate.
'
Public Property Let Capacity(ByVal Value As Long)
    If Value < mLength Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    If Value <> mCapacity Then
        If Value > 0 Then
            ReDim Preserve mString(0 To Value - 1)
        Else
            Erase mString
            SAPtr(mString) = SafeArrayCreateVector(vbString, 0, 0)
        End If
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
        Error.IndexOutOfRange ArgumentOutOfRange_Index
    
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
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    mString(Index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param requiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
'
Public Function EnsureCapacity(ByVal RequiredCapacity As Long) As Long
    Dim NewCapacity As Long
    
    If RequiredCapacity > mCapacity Then
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
    Equals = MyBase.Equals(Me, Value)
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
' @param Value The character as either a String or Integer value.
' @remarks The new value can be either String, in which case the
' first characters is used, or it can be a number.
'
Public Property Let FillCharacter(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbLong, vbInteger, vbByte
            mFillChar = AscW(ChrW$(Value))
        Case vbString
            mFillChar = AscW(Value)
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
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Removes all characters from the current <a href="StringBuilder.html">StringBuilder</a> instance.
'
' @return An object whose <a href="StringBuilder.get.Length.html">Length</a> is 0 (zero).
' @remarks <b>Clear</b> is a convenience method that is equivalent to setting the Length property of the current instance to 0 (zero).
'
Public Function Clear() As StringBuilder
    Length = 0
    Set Clear = Me
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
    Dim s As String
    Dim l As Long
    
    If Index < 0 Or Index > mLength Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    
    s = Object.ToString(Value)
    l = Len(s)
    
    If l > 0 And Count > 0 Then
        InsertSpace Index, l * Count
        
        Do While Count > 0
            CopyMemory mString(Index), ByVal StrPtr(s), l * vbSizeOfChar
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
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    If SAPtr(Chars) = vbNullPtr Then _
        Error.ArgumentNull "Chars", ArgumentNull_Array
    
    Dim Range As ListRange
    Range = MakeDefaultRange(StartIndex, LBound(Chars), CharCount, Len1D(Chars), Parameter_StartIndex, Parameter_CharCount)
    
    If Range.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "CharCount", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Chars) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    
    If Range.Count > 0 Then
        InsertSpace Index, Range.Count
        CopyMemory mString(Index), Chars(Range.Index), Range.Count * vbSizeOfChar
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
' @param Value The length of the string value.
' @remarks If the new value is longer than the current length, then the
' new space is padded with FillCharacter (default 32).
'
Public Property Let Length(ByVal Value As Long)
    If Value < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    
    If Value > mLength Then
        If Value > mCapacity Then
            EnsureCapacity Value
        End If
        
        Fill mLength, mFillChar, mCapacity - mLength
    End If
    
    mLength = Value
End Property

''
' Removes a number of characters from the string value.
'
' @param StartIndex The index in the string value to start removing from.
' @param Length The number of characters to remove from the string value.
' @return This instance of StringBuilder.
' @remarks StartIndex is zero-based.
'
Public Function Remove(ByVal StartIndex As Long, ByVal Length As Long) As StringBuilder
    If StartIndex < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_StartIndex
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    If StartIndex + Length > mLength Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Index
    
    If Length > 0 Then
        CopyMemory mString(StartIndex), mString(StartIndex + Length), (mLength - (StartIndex + Length)) * vbSizeOfChar
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
    Dim Range           As ListRange
    Dim MatchCount      As Long
    Dim MatchIndexes()  As Long
    Dim NewLength       As Long
    
    If Len(OldValue) = 0 Then _
        Error.Argument Argument_EmptyName, "OldValue"
    
    Range = MakeDefaultRange(StartIndex, 0, Count, mLength, Parameter_StartIndex)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_StartIndex
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > mLength Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Index
    
    MatchCount = FindMatches(OldValue, Range.Index, Range.Count, MatchIndexes)

    If MatchCount > 0 Then
        NewLength = mLength - ((Len(OldValue) - Len(NewValue)) * MatchCount)

        Select Case NewLength
            Case mLength
                ReplaceEqualLength NewValue, MatchIndexes, MatchCount
            Case Is < mLength
                ReplaceSmallerLength NewValue, MatchIndexes, MatchCount, Len(OldValue)
            Case Else
                ReplaceLargerLength NewValue, MatchIndexes, MatchCount, Len(OldValue), NewLength
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
    Dim Range As ListRange
    
    ' 99% of the calls will be to retrieve the entire string,
    ' so we do a quick check instead of going through the full
    ' optional argument gambit.
    If IsMissing(StartIndex) And IsMissing(Length) Then
        ToString = SysAllocStringLen(VarPtr(mString(0)), mLength)
        Exit Function
    End If
    
    Range = MakeDefaultRange(StartIndex, 0, Length, mLength, Parameter_StartIndex, Parameter_Length)
    
    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_StartIndex
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    If Range.Index + Range.Count > mLength Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_Index
    
    If Range.Count > 0 Then
        ToString = SysAllocStringLen(VarPtr(mString(Range.Index)), Range.Count)
    End If
End Function

Friend Sub InternalAppendFormat(ByVal Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant)
    Dim CustomFormatter     As ICustomFormatter
    Dim HasCustomFormatter  As Boolean
    Dim Index               As Long
    Dim Max                 As Long
    Dim SubStringStart      As Long
    Dim FormatChars()       As Integer
    
    ' See if the user wants to format anything their own way.
    If Not Provider Is Nothing Then
        Set CustomFormatter = Provider.GetFormat("ICustomFormatter")
        HasCustomFormatter = Not (CustomFormatter Is Nothing)
    End If
    
    Max = Len(Format)
    FormatChars = AllocChars(Format)
    
    On Error GoTo Catch
    
    ' Parse normal substring.
    Do While Index < Max
        Select Case FormatChars(Index)
            Case vbOpenBraceChar
                ' potential escape or format info.
                
                If Index > SubStringStart Then
                    ' If we have some characters, just append them now
                    ' even if the "{" may be escaped. Most times it won't be.
                    AppendCharsCore FormatChars, SubStringStart, Index - SubStringStart
                End If
                
                Index = Index + 1
                
                ' We test for escape by checking the next character for a "{".
                ' But if we're out of characters, then we ended with a "{",
                ' and that is an invalid format.
                If Index = Max Then FormatError
                
                ' We have enough characters to test for possible escape.
                If FormatChars(Index) = vbOpenBraceChar Then
                    ' escape it by setting the start of the substring to the second one.
                    SubStringStart = Index
                Else
                    Dim ArgIndex        As Long
                    Dim ArgWidth        As Long
                    Dim ArgFormatting   As String
                    
                    ' must be format info within a "{}", so call the parser for it.
                    Index = ParseArgumentFormatting(FormatChars, Index, ArgIndex, ArgWidth, ArgFormatting)
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
                        s = Object.ToString(Args(ArgIndex), ArgFormatting, Provider)
                    End If
                    
                    ' Add leading padding.
                    If ArgWidth > Len(s) Then
                        AppendChar vbSpaceChar, ArgWidth - Len(s)
                    End If
                    
                    ' Finally append the formatted value to the string.
                    AppendString s
                    
                    ' Add trailing padding.
                    If -ArgWidth > Len(s) Then
                        AppendChar vbSpaceChar, (-ArgWidth) - Len(s)
                    End If
                    
                    SubStringStart = Index + 1
                    
                End If
            Case vbCloseBraceChar
                ' must be escaped or else it is a format error.
                
                Index = Index + 1
                
                ' Need enough characters to test for escape.
                If Index = Max Then FormatError
                
                If FormatChars(Index) = vbCloseBraceChar Then
                    ' escape it and add the substring to this point.
                    ' The substring ends one character back because we
                    ' don't want to include both "}}" in the substring.
                    If Index - 1 > SubStringStart Then
                        ' Append the current substring.
                        AppendCharsCore FormatChars, SubStringStart, Index - SubStringStart - 1
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
        AppendCharsCore FormatChars, SubStringStart, Index - SubStringStart
    End If
    
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars FormatChars
    Return
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByRef StartIndex As Variant, ByRef Length As Variant, ByVal Capacity As Long)
    Dim ElemIndex As Long
    Dim ElemCount As Long
    ElemIndex = CLngOrDefault(StartIndex, 0)
    ElemCount = CLngOrDefault(Length, Len(s) - ElemIndex)
        
    If ElemIndex < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_StartIndex
    If ElemCount < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NeedNonNegNum
    If ElemIndex + ElemCount > Len(s) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_IndexLength
    If Capacity < 0 Then _
        Error.ArgumentOutOfRange "Capacity", ArgumentOutOfRange_NeedNonNegNum
    
    If Capacity = 0 Then
        Capacity = DefaultCapacity
    End If
    
    If ElemCount > Capacity Then
        Capacity = ElemCount
    End If
    
    If Capacity <> mCapacity Then
        Me.Capacity = Capacity
    End If
      
    If ElemCount > 0 Then
        CopyMemory mString(0), ByVal StrPtr(s) + ElemIndex * vbSizeOfChar, ElemCount * vbSizeOfChar
        mLength = ElemCount
    End If
End Sub

Private Sub Class_Initialize()
    mFillChar = DefaultFillChar
    mCapacity = DefaultCapacity
    ReDim mString(0 To DefaultCapacity - 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AppendCharsValue(ByRef Value As Variant, ByRef StartIndex As Variant, ByRef CharCount As Variant)
    Dim Chars() As Integer
    Dim Range As ListRange
    
    ValidateArray Value, Parameter_Value
    Chars = AsChars(Value)
    
    On Error GoTo Catch
    Range = MakeDefaultRange(StartIndex, LBound(Value), CharCount, Len1D(Value), Parameter_StartIndex)
    
    If Range.Index < LBound(Value) Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count - 1 > UBound(Value) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_IndexCount

    AppendCharsCore Chars, Range.Index, Range.Count
    
    FreeChars Chars
    Exit Sub

Catch:
    FreeChars Chars
    Throw
End Sub

Private Sub AppendStringValue(ByRef Value As Variant, ByRef StartIndex As Variant, ByRef Count As Variant)
    Dim ValuePtr    As Long
    Dim ValueLength As Long

    ValuePtr = StrPtr(Value)
    ValueLength = Len(Value)

    Dim Range As ListRange
    Range = MakeDefaultRange(StartIndex, 0, Count, ValueLength, Parameter_StartIndex)

    If Range.Index < 0 Then _
        Error.ArgumentOutOfRange "StartIndex", ArgumentOutOfRange_Index
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Range.Index + Range.Count > ValueLength Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_IndexCount

    AppendCore ValuePtr, Range.Index, Range.Count
End Sub

Private Sub AppendCharsCore(ByRef Chars() As Integer, ByVal StartIndex As Long, ByVal CharCount As Long)
    Debug.Assert SAPtr(Chars) <> vbNullPtr
    
    If CharCount > 0 Then
        Dim FirstElementPtr As Long
        FirstElementPtr = VarPtr(Chars(LBound(Chars)))
        
        AppendCore FirstElementPtr, StartIndex, CharCount
    End If
End Sub

Private Sub AppendCore(ByVal CharPtr As Long, ByVal StartIndex As Long, ByVal CharCount As Long)
    If CharCount > 0 Then
        Dim NewLength As Long
        NewLength = mLength + CharCount
        EnsureCapacity NewLength
        CopyMemory mString(mLength), ByVal CharPtr + StartIndex * vbSizeOfChar, CharCount * vbSizeOfChar
        mLength = NewLength
    End If
End Sub

Private Sub Fill(ByVal Index As Long, ByVal Char As Integer, ByVal Count As Long)
    Dim i As Long
    
    For i = Index To Index + Count - 1
        mString(i) = Char
    Next i
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidString))
End Sub

Private Sub InsertSpace(ByVal Index As Long, ByVal Size As Long)
    Dim NewLength As Long
    NewLength = mLength + Size
    
    EnsureCapacity NewLength
    
    If Index < mLength Then
        ' Create space in the string. The new space is uninitialized.
        CopyMemory mString(Index + Size), mString(Index), (mLength - Index) * vbSizeOfChar
    End If
    
    mLength = NewLength
End Sub

''
' We must be inside "{}". We must end with a matching "}" or it is a format error.
'
' @param Index The index of the first character after the "{".
' @param ArgIndex The index of the argument in the Args array.
' @param ArgWidth The width of the column the argument will be displayed in.
' @param ArgFormat The format information used to determine formatting for the argument.
' @return The index of the terminating "}".
'
Private Function ParseArgumentFormatting(ByRef FormatChars() As Integer, ByVal Index As Long, ByRef ArgIndex As Long, ByRef ArgWidth As Long, ByRef ArgFormat As String) As Long
    Dim Max As Long
    
    Max = UBound(FormatChars) + 1
    
    ' Get the argument index.
    ArgIndex = ParseNumber(FormatChars, Index)
    Index = ConsumeSpaces(FormatChars, Index)
    
    If FormatChars(Index) = vbCommaChar Then
        Index = ConsumeSpaces(FormatChars, Index + 1)
        
        ' Get the column width and aligment.
        ' Negative values mean align left in the column.
        ArgWidth = ParseNumber(FormatChars, Index)
        Index = ConsumeSpaces(FormatChars, Index)
    Else
        ' Make sure they know there was no width specified.
        ArgWidth = 0
    End If
    
    ' Check if formatting is being specified.
    If FormatChars(Index) = vbColonChar Then
        Dim StartIndex As Long
        
        ' everything from here to the "}" is part of the argument formatting.
        Index = Index + 1
        StartIndex = Index
        Do
            ' can never reach the end until after a "}".
            If Index = Max Then FormatError
            If FormatChars(Index) = vbCloseBraceChar Then
                ArgFormat = SysAllocStringLen(VarPtr(FormatChars(StartIndex)), Index - StartIndex)
                Exit Do
            End If
            Index = Index + 1
        Loop
    Else
        ' Make sure we let them know that no formatting was supplied.
        ArgFormat = vbNullString
    End If
    
    ' if we get here we should have reached the closing "}".
    If FormatChars(Index) <> vbCloseBraceChar Then FormatError
    
    ParseArgumentFormatting = Index
End Function

''
' Advances the index to the next non-space character.
'
' @param The index to the first space to be consumed.
' @return The index of the first character after the contiguous
' stream of spaces have been consumed.
'
Private Function ConsumeSpaces(ByRef Chars() As Integer, ByVal Index As Long) As Long
    Dim MaxIndex As Long
    
    MaxIndex = UBound(Chars)
    
    Do
        ' Can never hit the end consuming spaces.
        If Index > MaxIndex Then FormatError
        
        If Chars(Index) <> vbSpaceChar Then
            Exit Do
        End If
        
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
Private Function ParseNumber(ByRef Chars() As Integer, ByRef Index As Long) As Long
    Dim MaxIndex As Long
    Dim Sign As Long
    Dim FirstCharIndex As Long
    Dim Number      As Long
    Dim NumberFound As Boolean
    
    MaxIndex = UBound(Chars)
    Sign = 1
    FirstCharIndex = Index
    
    Do
        ' We can never hit the end parsing a number.
        If Index > MaxIndex Then FormatError
        
        Dim Ch As Integer
        Ch = Chars(Index)
        
        Select Case Ch
            Case vbZeroChar To vbNineChar
                ' Keep adding digits to the number
                Number = Number * 10 + Ch - vbZeroChar
                NumberFound = True
            Case vbMinusChar
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
    Dim MatchCount      As Long
    Dim OldValueLength  As Long
    Dim OldChars()      As Integer
    Dim FirstOldChar    As Integer
    Dim Matched         As Boolean
    
    ReDim MatchIndexes(0 To 4)
    OldValueLength = Len(OldValue)
    OldChars = AllocChars(OldValue)
    
    ' We'll cache this since we'll be hitting it constantly.
    FirstOldChar = OldChars(0)
    
    ' Don't need to search at the end if the old value
    ' can't fit. If the end matches the old value, then
    ' the inner loop will finish out the matching.
    Count = Count - OldValueLength + 1
        
    ' Loop through the current StringBuilder string trying to find
    ' all the matches against the old value to be replaced. If a
    ' match is found, then the index to that match is stored for
    ' future usage when replacing the matched old values with the new.
    Do While Count > 0
        ' Found the start of a possible match.
        If mString(StartIndex) = FirstOldChar Then
            ' We'll be opitmistic and assume a match.
            Matched = True
            
            ' Start with the second character.
            Dim i As Long
            For i = 1 To OldValueLength - 1
                If mString(StartIndex + i) <> OldChars(i) Then
                    Matched = False
                    Exit For
                End If
            Next
            
            If Matched Then
                If MatchCount > UBound(MatchIndexes) Then
                    ' like .NET we'll grow faster initially, but slow down later.
                    ReDim Preserve MatchIndexes(0 To (UBound(MatchIndexes) + 1) * 3 \ 2 + 4)
                End If
                
                MatchIndexes(MatchCount) = StartIndex
                MatchCount = MatchCount + 1
                StartIndex = StartIndex + OldValueLength - 1
                Count = Count - OldValueLength + 1
            End If
        End If
        
        StartIndex = StartIndex + 1
        Count = Count - 1
    Loop

    FreeChars OldChars
    FindMatches = MatchCount
End Function

''
' Optimized to simply copy the new value over the old value within the original buffer.
'
' @remarks This is used when the old and new values are the same length, resulting in
' no change in the size of the final string.
'
Private Sub ReplaceEqualLength(ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long)
    Dim NewValuePtr     As Long
    Dim AmountToCopy    As Long
    
    NewValuePtr = StrPtr(NewValue)
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
    ReplaceMatches mString, mString, NewValue, MatchIndexes, MatchCount, OldValueLength
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
    Dim NewString() As Integer
    
    NewCapacity = CorMath.Max(NewLength, mCapacity)
    ReDim NewString(0 To NewCapacity - 1)
    ReplaceMatches mString, NewString, NewValue, MatchIndexes, MatchCount, OldValueLength
    Helper.Swap4 ByVal ArrPtr(mString), ByVal ArrPtr(NewString)
    mCapacity = NewCapacity
End Sub

''
' Performs the replacement of old values with new values, copying the unaltered text from the
' source buffer and placing the final unaltered text plus newly replaced text in the destination buffer.
'
Private Sub ReplaceMatches(ByRef SourceBuffer() As Integer, ByRef DestBuffer() As Integer, ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long, ByVal OldValueLength As Long)
    Dim NewValuePtr As Long
    Dim NewValueLen As Long
    Dim DestIndex   As Long
    
    NewValuePtr = StrPtr(NewValue)
    NewValueLen = Len(NewValue)
    
    Dim OriginalIndex   As Long
    Dim AmountToCopy    As Long
    Dim i               As Long
    For i = 0 To MatchCount - 1
        AmountToCopy = MatchIndexes(i) - OriginalIndex
        
        If AmountToCopy > 0 Then
            CopyMemory DestBuffer(DestIndex), SourceBuffer(OriginalIndex), AmountToCopy * vbSizeOfChar
            DestIndex = DestIndex + AmountToCopy
        End If
        
        If NewValueLen > 0 Then
            CopyMemory DestBuffer(DestIndex), ByVal NewValuePtr, NewValueLen * vbSizeOfChar
            DestIndex = DestIndex + NewValueLen
        End If
        
        OriginalIndex = OriginalIndex + AmountToCopy + OldValueLength
    Next i
    
    If OriginalIndex < mLength Then
        CopyMemory DestBuffer(DestIndex), SourceBuffer(OriginalIndex), (mLength - OriginalIndex) * vbSizeOfChar
    End If
End Sub

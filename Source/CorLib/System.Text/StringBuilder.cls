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
Private mStringPtr  As Long
Private mCapacity   As Long
Private mLength     As Long
Private mFillChar   As Integer


''
' Appends the string representation of the specified value to this instance.
'
' @param Value The value to be appended.
' @param StartIndex A zero-based index of the first character in the string representation to begin appending from.
' @param Count The number of characters from the string representation to append.
' @return A refernece to this instance after the append operation has completed.
' @remarks This version of <b>Append(Value)</b> attempts to convert to the value using <h href="ObjectStatic.html">Object.ToString(Value)</a>.
' This will allow for objects that implement <a href="IObject.html">IObject</a> or <a href="IFormattable.html">IFormattable</a> to be converted to a string value using additional information.
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
' @remarks <p>The <b>AppendChar(Char)</b> method modifies the existing instance of this class; it does not return a new class instance.</p>
' <p>The <i>Char</i> argument accepts a <b>Long</b> datatype, however, it converts the value to an <b>Integer</b> by taking only the lower 16 bits.</p>
Public Function AppendChar(ByVal Char As Long, Optional ByVal Count As Long = 1) As StringBuilder
    If Count < 0 Then _
        Error.ArgumentOutOfRange "Count", ArgumentOutOfRange_NeedNonNegNum
    If Char < MinCharValue Or Char > MaxCharValue Then _
        Error.Overflow Overflow_Char
        
    EnsureCapacity mLength + Count
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If Count = 1 Then
#If Release Then
        mString(mLength) = Char And &HFFFF&
#Else
        mString(mLength) = AsWord(Char)
#End If
    Else
#If Release Then
        Fill mLength, Char And &HFFFF&, Count
#Else
        Fill mLength, AsWord(Char), Count
#End If
    End If
    
    mLength = mLength + Count
    Set AppendChar = Me
End Function

''
' Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument.
'
' @param Format The composite format string containing formatting information.
' @param Args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @see NumberFormatInfo
' @see DateTimeFormatInfo
' @include "..\DocIncludes\StringBuilder.AppendFormat.html"
Public Function AppendFormat(ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    Set AppendFormat = AppendFormatArrayEx(Nothing, Format, SwappedArgs)
End Function

''
' Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument.
'
' @param Provider A custom formatting object that formats the arguments.
' @param Format The composite format string containing formatting information.
' @param Args The arguments to be formatted into the string.
' @return This instance of StringBuilder.
' @remarks <p>This method uses the composite formatting feature of VBCorLib to convert the value of an object to its text representation and embed that representation in the current <a href="StringBuilder.html">StringBuilder</a> object.</p>
' <p>The <i>Format</i> parameter consists of zero or more runs of text intermixed with zero or more indexed placeholders, called format items, that correspond to values in the parameter list of this method. The formatting process replaces each format item with the string representation of the corresponding value.</p>
' <p>The syntax of a format item is as follows:</p>
' <p>{<i>index</i>[,<i>length</i>][:<i>formatString</i>]}</p>
' <p>Elements in square brackets are optional. The following table describes each element.</p>
' <table>
'    <tr>
'        <th>Element</th><th>Description</th>
'    </tr>
'    <tr>
'        <td><i>index</i></td><td>The zero-based position in the parameter list of the object to be formatted. If the object specified by <i>index</i> is <b>Nothing</b>, the format item is replaced by an empty string. If there is no parameter in the <i>index</i> position, a <a href="FormatException.html">FormatException</a> is thrown.</td>
'    </tr>
'    <tr>
'        <td>,<i>length</i></td><td>The minimum number of characters in the string representation of the parameter. If positive, the parameter is right-aligned; if negative, it is left-aligned.</td>
'    </tr>
'    <tr>
'        <td>:<i>formatString</i></td><td>A standard or custom format string that is supported by the parameter.</td>
'    </tr>
' </table>
' <p>The provider parameter specifies an <a href="IFormatProvider.html">IFormatProvider</a> implementation that can provide formatting information for the objects in <i>args</i>. provider can be any of the following:</p>
' <ul>
'       <li>A <a href="CultureInfo.html">CultureInfo</a> object that provides culture-specific formatting information.</li>
'       <li>A <a href="NumberFormatInfo.html">NumberFormatInfo</a> object that provides culture-specific formatting information for numeric values in <i>Args</i>.</li>
'       <li>A <a href="DateTimeFormatInfo.html">DateTimeFormatInfo</a> object that provides culture-specific formatting information for date and time values in <i>Args</i>.</li>
'       <li>A custom <a href="IFormatProvider.html">IFormatProvider</a> implementation that provides formatting information for one or more of the objects in <i>Args</i>. Typically, such an implementation also implements the <a href="ICustomFormatter.html">ICustomFormatter</a> interface.</li>
' </ul>
' <p>If the <i>Provider</i> parameter is <b>Nothing</b>, format provider information is obtained from the current culture.</p>
' <p><i>Args</i> represents the objects to be formatted. Each format item in <i>Format</i> is replaced with the string representation of the corresponding object in <i>args</i>. If the format item includes <i>formatString</i> and the corresponding object in <i>Args</i> implements the <a href="IFormattable.html">IFormattable</a> interface, then <b>Args(index).ToString(formatString, provider)</b> defines the formatting. Otherwise, <b>Object.ToString(Args(index))</b> defines the formatting.</p>
' @see NumberFormatInfo
' @see DateTimeFormatInfo
Public Function AppendFormatEx(ByRef Provider As IFormatProvider, ByRef Format As String, ParamArray Args() As Variant) As StringBuilder
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(20)
    Set AppendFormatEx = AppendFormatArrayEx(Provider, Format, SwappedArgs)
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the <a href="StringBuilder.html">StringBuilder</a>.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of <a href="StringBuilder.html">StringBuilder</a>.
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
' @remarks This method appends the entire string to the end of the <a href="StringBuilder.html">StringBuilder</a>. There is no instance returned.
Public Sub AppendString(ByRef s As String)
'    AppendCore StrPtr(s), 0, Len(s)
    Dim Length As Long
    
    Length = Len(s)
    
    If Length Then
        If mLength + Length > mCapacity Then
            EnsureCapacity mLength + Length
        End If
        
        CopyMemory ByVal mStringPtr + mLength * vbSizeOfChar, ByVal StrPtr(s), Length * vbSizeOfChar
        mLength = mLength + Length
    End If
End Sub

''
' Copies a set of characters from the <a href="StringBuilder.html">StringBuilder</a> into a character array.
'
' @param SourceIndex The starting index in the <a href="StringBuilder.html">StringBuilder</a> to be copying from (0-based).
' @param Destination The character array to copy to.
' @param DestinationIndex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
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
' @remarks The capacity of the current instance is how much memory is currently allocated for containing
' the current string representation. It is not the length of the string itself.
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Sets the amount of memory allocated in characters.
'
' @param Value The number of characters to allocate.
' @remarks The capacity of the current instance is how much memory is currently allocated for containing
' the current string representation. It is not the length of the string itself.
Public Property Let Capacity(ByVal Value As Long)
    If Value < mLength Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_SmallCapacity
    
    If Value <> mCapacity Then
        If Value > 0 Then
            ReDim Preserve mString(0 To Value - 1)
            mStringPtr = VarPtr(mString(0))
        Else
            Erase mString
            SAPtr(mString) = SafeArrayCreateVector(vbString, 0, 0)
            mStringPtr = vbNullPtr
        End If
        
        mCapacity = Value
    End If
End Property

''
' Returns a character from the string value.
'
' @param Index The index from the string value to return the character.
' @return The character in the string at the specified index.
' @remarks <i>Index</i> is zero-base.
Public Property Get Chars(ByVal Index As Long) As Integer
Attribute Chars.VB_UserMemId = 0
    If Index < 0 Or Index >= mLength Then _
        Error.IndexOutOfRange ArgumentOutOfRange_Index
    
    Chars = mString(Index)
End Property

''
' Sets a character in the string value.
'
' @param Index The index into the string value to set the character.
' @param Char The character to set.
' @remarks <p><i>Index</i> is zero-based.</p>
' <p>Setting a character can only replace an existing character. A character cannot be appended using this method.
' To append a character use <a href="StringBuilder.AppendChar.html">StringBuilder.AppendChar</a>.</p>
' @see #AppendChar
Public Property Let Chars(ByVal Index As Long, ByVal Char As Integer)
    If Index < 0 Or Index >= mLength Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_Index
    
    mString(Index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param RequiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
' @remarks When <i>RequiredCapacity</i> is larger than the current capacity, The new capacity is
' calculated as double the current capacity. If the new capacity is still smaller than <i>RequiredCapacity</i>, <i>RequiredCapacity</i> becomes the new capacity.
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
' Returns a boolean indicating if the value and this object instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @remarks This implementation simply delegates to <a href="ObjectStatic.Equals.html">Object.Equals</a>.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns the fill character used to pad empty space.
'
' @return An <b>Integer</b> representing the <b>AscW</b> of the character.
' @remarks The return type is a <b>Variant</b> with a sub-type <i>Integer</i>.
Public Property Get FillCharacter() As Variant
    FillCharacter = mFillChar
End Property

''
' Sets the fill character used to pad empty space.
'
' @param Value The character used to as padding of empty space.
' @remarks <p>The property accepts <b>String</b>, <b>Long</b>, <b>Integer</b>, and <b>Byte</b> datatypes.</p>
' <p>The datatypes are converted using the following methods:</p>
' <table style="width: auto">
'   <tr>
'       <th>Datatype</th>
'       <th>Method</th>
'   </tr>
'   <tr>
'        <td><b>String</b></td><td><b>AscW(</b><i>Value</i><b>)</b></td>
'   </tr>
'   <tr>
'        <td><b>Long</b>,<b>Integer</b>,<b>Byte</b></td><td><b>AscW(ChrW$(</b><i>Value</i><b>))</b></td>
'   </tr>
' </table>
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
' @remarks This implementation simply delegates to <a href="ObjectStatic.GetHashCode.html">Object.GetHashCode</a>.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Removes all characters from the current <a href="StringBuilder.html">StringBuilder</a> instance.
'
' @return An object whose <a href="StringBuilder.get.Length.html">Length</a> is 0 (zero).
' @remarks <b>Clear</b> is a convenience method that is equivalent to setting the Length property of the current instance to 0 (zero).
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
' @return This instance of <a href="StringBuilder.html">StringBuilder</a>.
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
Public Property Get Length() As Long
    Length = mLength
End Property

''
' Sets the length of the current string value.
'
' @param Value The length of the string value.
' @remarks If the new value is longer than the current length, then the
' new space is padded with <a href="StringBuilder.set.FillCharacter.html">FillCharacter</a> (default 32).
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
' @return This instance of <a href="StringBuilder.html">StringBuilder</a>.
' @remarks StartIndex is zero-based.
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
' @param StartIndex The index of the start of the substring in the <a href="StringBuilder.html">StringBuilder</a> object.
' @param count The number of characters in the <a href="StringBuilder.html">StringBuilder</a> object substring.
' @return This instance of <a href="StringBuilder.html">StringBuilder</a>.
' @remarks Use <i>StartIndex</i> and <i>Count</i> to replace the old value with a substring of the new value. <i>StartIndex</i> is zero-based.
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
' @param StartIndex The index to start the return substring from.
' @param Length The number of characters to return in the string.
' @return A string or substring representing the internal string of <a href="StringBuilder.html">StringBuilder</a>.
' @remarks <i>StartIndex</i> is zero-based.
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

''
' Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array.
' @param Format A composite format string (see Remarks).
' @param Args An array of values to format.
' @return A reference to this instance after the append operation has completed. After the append operation, this instance contains any data that existed before the operation, suffixed by a copy of <i>Format</i> where any format specification is replaced by the string representation of the corresponding object argument.
' @remarks <p>This method uses the composite formatting feature of VBCorLib to convert the value of an object to its text representation and embed that representation in the current <a href="StringBuilder.html">StringBuilder</a> object.</p>
' <p>The <i>Format</i> parameter consists of zero or more runs of text intermixed with zero or more indexed placeholders, called format items, that correspond to values in the parameter list of this method. The formatting process replaces each format item with the string representation of the corresponding value.</p>
' <p>The syntax of a format item is as follows:</p>
' <p>{<i>index</i>[,<i>length</i>][:<i>formatString</i>]}</p>
' <p>Elements in square brackets are optional. The following table describes each element.</p>
' <table>
'    <tr>
'        <th>Element</th><th>Description</th>
'    </tr>
'    <tr>
'        <td><i>index</i></td><td>The zero-based position in the parameter list of the object to be formatted. If the object specified by <i>index</i> is <b>Nothing</b>, the format item is replaced by an empty string. If there is no parameter in the <i>index</i> position, a <a href="FormatException.html">FormatException</a> is thrown.</td>
'    </tr>
'    <tr>
'        <td>,<i>length</i></td><td>The minimum number of characters in the string representation of the parameter. If positive, the parameter is right-aligned; if negative, it is left-aligned.</td>
'    </tr>
'    <tr>
'        <td>:<i>formatString</i></td><td>A standard or custom format string that is supported by the parameter.</td>
'    </tr>
' </table>
' <p><i>Args</i> represents the objects to be formatted. Each format item in <i>Format</i> is replaced with the string representation of the corresponding object in <i>args</i>. If the format item includes <i>formatString</i> and the corresponding object in <i>Args</i> implements the <a href="IFormattable.html">IFormattable</a> interface, then <b>Args(index).ToString(formatString, provider)</b> defines the formatting. Otherwise, <b>Object.ToString(Args(index))</b> defines the formatting.</p>
Public Function AppendFormatArray(ByRef Format As String, ByRef Args() As Variant) As StringBuilder
    Set AppendFormatArray = AppendFormatArrayEx(Nothing, Format, Args)
End Function

''
' Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array using a specified format provider.
'
' @param Provider An object that supplies culture-specific formatting information.
' @param Format A composite format string (see Remarks).
' @param Args An array of values to format.
' @return A reference to this instance after the append operation has completed. After the append operation, this instance contains any data that existed before the operation, suffixed by a copy of <i>Format</i> where any format specification is replaced by the string representation of the corresponding object argument.
' @remarks <p>This method uses the composite formatting feature of VBCorLib to convert the value of an object to its text representation and embed that representation in the current <a href="StringBuilder.html">StringBuilder</a> object.</p>
' <p>The <i>Format</i> parameter consists of zero or more runs of text intermixed with zero or more indexed placeholders, called format items, that correspond to values in the parameter list of this method. The formatting process replaces each format item with the string representation of the corresponding value.</p>
' <p>The syntax of a format item is as follows:</p>
' <p>{<i>index</i>[,<i>length</i>][:<i>formatString</i>]}</p>
' <p>Elements in square brackets are optional. The following table describes each element.</p>
' <table>
'    <tr>
'        <th>Element</th><th>Description</th>
'    </tr>
'    <tr>
'        <td><i>index</i></td><td>The zero-based position in the parameter list of the object to be formatted. If the object specified by <i>index</i> is <b>Nothing</b>, the format item is replaced by an empty string. If there is no parameter in the <i>index</i> position, a <a href="FormatException.html">FormatException</a> is thrown.</td>
'    </tr>
'    <tr>
'        <td>,<i>length</i></td><td>The minimum number of characters in the string representation of the parameter. If positive, the parameter is right-aligned; if negative, it is left-aligned.</td>
'    </tr>
'    <tr>
'        <td>:<i>formatString</i></td><td>A standard or custom format string that is supported by the parameter.</td>
'    </tr>
' </table>
' <p>The provider parameter specifies an <a href="IFormatProvider.html">IFormatProvider</a> implementation that can provide formatting information for the objects in <i>args</i>. provider can be any of the following:</p>
' <ul>
'       <li>A <a href="CultureInfo.html">CultureInfo</a> object that provides culture-specific formatting information.</li>
'       <li>A <a href="NumberFormatInfo.html">NumberFormatInfo</a> object that provides culture-specific formatting information for numeric values in <i>Args</i>.</li>
'       <li>A <a href="DateTimeFormatInfo.html">DateTimeFormatInfo</a> object that provides culture-specific formatting information for date and time values in <i>Args</i>.</li>
'       <li>A custom <a href="IFormatProvider.html">IFormatProvider</a> implementation that provides formatting information for one or more of the objects in <i>Args</i>. Typically, such an implementation also implements the <a href="ICustomFormatter.html">ICustomFormatter</a> interface.</li>
' </ul>
' <p>If the <i>Provider</i> parameter is <b>Nothing</b>, format provider information is obtained from the current culture.</p>
' <p><i>Args</i> represents the objects to be formatted. Each format item in <i>Format</i> is replaced with the string representation of the corresponding object in <i>args</i>. If the format item includes <i>formatString</i> and the corresponding object in <i>Args</i> implements the <a href="IFormattable.html">IFormattable</a> interface, then <b>Args(index).ToString(formatString, provider)</b> defines the formatting. Otherwise, <b>Object.ToString(Args(index))</b> defines the formatting.</p>
Public Function AppendFormatArrayEx(ByVal Provider As IFormatProvider, ByRef Format As String, ByRef Args() As Variant) As StringBuilder
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
                        Error.Format Format_IndexOutOfRange
                        
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
    
    Set AppendFormatArrayEx = Me
    FreeChars FormatChars
    Exit Function
    
Catch:
    FreeChars FormatChars
    ThrowOrErr Err
End Function


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
    mStringPtr = VarPtr(mString(0))
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
    ThrowOrErr Err
End Sub

Private Sub AppendStringValue(ByRef Value As Variant, ByRef StartIndex As Variant, ByRef Count As Variant)
    Dim ValuePtr    As Long
    Dim ValueLength As Long
    Dim Range       As ListRange

    ValuePtr = StrPtr(Value)
    ValueLength = Len(Value)
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
    Dim FirstElementPtr As Long
    
    Debug.Assert SAPtr(Chars) <> vbNullPtr
    
    If CharCount > 0 Then
        FirstElementPtr = VarPtr(Chars(LBound(Chars)))
        AppendCore FirstElementPtr, StartIndex, CharCount
    End If
End Sub

Private Sub AppendCore(ByVal CharPtr As Long, ByVal StartIndex As Long, ByVal CharCount As Long)
    Dim NewLength As Long
    
    If CharCount > 0 Then
        NewLength = mLength + CharCount
        
        If NewLength > mCapacity Then
            EnsureCapacity NewLength
        End If
        
        CopyMemory ByVal mStringPtr + mLength * vbSizeOfChar, ByVal CharPtr + StartIndex * vbSizeOfChar, CharCount * vbSizeOfChar
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

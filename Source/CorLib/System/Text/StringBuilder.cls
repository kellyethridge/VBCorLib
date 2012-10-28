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
'    Module: StringBuilder
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

Private Const DEF_FILLCHAR          As Integer = vbSpace
Private Const DEF_CAPACITY          As Long = 512
Private Const OPEN_BRACE            As Long = 123
Private Const CLOSE_BRACE           As Long = 125

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
Public Function Append(ByRef value As Variant, Optional ByRef startIndex As Variant, Optional ByRef count As Variant) As StringBuilder
    Set Append = AppendString(Convert.ToString(value), startIndex, count)
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
Public Function AppendChar(ByVal Char As Long, Optional ByVal count As Long = 1) As StringBuilder
    If count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Count cannot be negative.", "Count", count)
    If Char < -&H7FFF& Or Char > &HFFFF& Then _
        Throw Cor.NewArgumentOutOfRangeException("Char must be between -32767 and 65535 inclusively.", "Char", Char)
    
    If Char > &H7FFF& Then Char = Char - &H10000
    
    If mLength + count > mCapacity Then Call EnsureCapacity(mLength + count)
    
    ' appending 1 character at a time happens 99%
    ' of the time, so optimize for it.
    If count = 1 Then
        mString(mLength) = Char
    Else
        Call Fill(mLength, Char, count)
    End If
    mLength = mLength + count
    Set AppendChar = Me
End Function

''
' Appends an array of characters to the end of the string value.
'
' @param Chars The array of characters to be appended.
' @param startindex The index from Chars to start appending.
' @param count The number of characters to append.
' @return This instance of StringBuilder.
'
Public Function AppendChars(ByRef Chars() As Integer, Optional ByRef startIndex As Variant, Optional ByRef count As Variant) As StringBuilder
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    
    result = GetOptionalArrayRange(SAPtr(Chars), startIndex, ElemIndex, count, ElemCount)
    If result <> NO_ERROR Then Call ThrowArrayRangeException(result, "Chars", ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex))
    
    If ElemCount = 0 Then
        Set AppendChars = Me
        Exit Function
    End If
    
    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(mString(mLength), Chars(ElemIndex), ElemCount * 2)
    mLength = mLength + ElemCount
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
Public Function AppendFormat(ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(16))
    Call InternalAppendFormat(Nothing, Format, vArgs)
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
Public Function AppendFormatEx(ByVal Provider As IFormatProvider, ByRef Format As String, ParamArray args() As Variant) As StringBuilder
    Dim vArgs() As Variant
    Call Helper.Swap4(ByVal ArrPtr(vArgs), ByVal Helper.DerefEBP(20))
    Call InternalAppendFormat(Provider, Format, vArgs)
    Set AppendFormatEx = Me
End Function

''
' Appends a string to the current string value.
'
' @param s String to be appended.
' @param StartIndex A zero-based index of the first character to start appending from in the string.
' @param Count The number of characters from the string to append.
' @return A reference to this instance of StringBuilder.
' @remarks<p>This version of append is designed for fast concatentation
' of the supplied string value. The entire string is appended.
' <p>This function does not return a reference to the instance of StringBuilder
' in order to keep the overhead cost down on the function call.</p>
'
Public Function AppendString(ByRef s As String, Optional ByRef startIndex As Variant, Optional ByRef count As Variant) As StringBuilder
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim result      As Long
    result = GetOptionalListRange(Len(s), startIndex, ElemIndex, count, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex))

    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(ByVal mStringPtr + mLength * 2, ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
    mLength = mLength + ElemCount
    Set AppendString = Me
End Function

''
' Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.
'
' @param s The string to append to the builder, followed by a Carraige-Return/Linefeed.
' @return A reference to this instance of StringBuilder.
'
Public Function AppendLine(Optional ByVal s As String) As StringBuilder
    Call AppendQuick(s & vbCrLf)
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
    Dim ElemCount As Long
    ElemCount = Len(s)
    If ElemCount = 0 Then Exit Sub
    If mLength + ElemCount > mCapacity Then Call EnsureCapacity(mLength + ElemCount)
    Call CopyMemory(ByVal mStringPtr + mLength * 2, ByVal StrPtr(s), ElemCount * 2)
    mLength = mLength + ElemCount
End Sub

''
' Copies a set of characters from the StringBuilder into a character array.
'
' @param sourceindex The starting index in the StringBuilder to be copying from (0-based).
' @param Destination The character array to copy to.
' @param destinationindex The start index to begin placing characters in to.
' @param Count The number of characters to copy.
'
Public Sub CopyTo(ByVal SourceIndex As Long, ByRef Destination() As Integer, ByVal DestinationIndex As Long, ByVal count As Long)
    Dim result As Long
    result = VerifyArrayRange(SAPtr(Destination), DestinationIndex, count)
    If result <> NO_ERROR Then _
        ThrowArrayRangeException result, "Destination", DestinationIndex, "DestinationIndex", count, "Count"
    
    result = VerifyListRange(mLength, SourceIndex, count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, SourceIndex, "SourceIndex", count, "Count")
    
    If count > 0 Then Call CopyMemory(Destination(DestinationIndex), ByVal mStringPtr + SourceIndex * 2, count * 2)
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
    
    Dim Ptr As Long
    Ptr = CoTaskMemAlloc(RHS * 2)
    If Ptr = vbNullPtr Then Throw New OutOfMemoryException
    Call CopyMemory(ByVal Ptr, ByVal mStringPtr, mLength * 2)
    Call CoTaskMemFree(mStringPtr)
    mStringPtr = Ptr
    mStringSA.pvData = Ptr
    mCapacity = RHS
End Property

''
' Returns a character from the string value.
'
' @param Index The index from the string value to return the character.
' @return The character in the string at the specified index.
' @remarks Index is zero-base.
'
Public Property Get Chars(ByVal index As Long) As Integer
Attribute Chars.VB_UserMemId = 0
    If index < 0 Or index >= mLength Then _
        Throw Cor.NewIndexOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Chars = mString(index)
End Property

''
' Sets a character in the string value.
'
' @param index The index into the string value to set the character.
' @param char The character to set.
' @remarks index is zero-based.
'
Public Property Let Chars(ByVal index As Long, ByVal Char As Integer)
    If index < 0 Or index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_Index), "index", index)
    
    mString(index) = Char
End Property

''
' Ensures that the internal buffer has atleast the requested amount.
'
' @param requiredCapacity The minimum amount of characters to be allocated.
' @return The capacity after the call is complete.
'
Public Function EnsureCapacity(ByVal RequiredCapacity As Long) As Long
    If RequiredCapacity <= mCapacity Then
        EnsureCapacity = mCapacity
        Exit Function
    End If
    
    Dim NewCapacity As Long
    NewCapacity = mCapacity * 2
    If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
    Capacity = NewCapacity
    EnsureCapacity = NewCapacity
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef value As Variant) As Boolean
    Equals = Object.Equals(Me, value)
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
Public Function Insert(ByVal index As Long, ByVal value As String, Optional ByVal count As Long = 1) As StringBuilder
    If index < 0 Or index > mLength Then _
        Throw Cor.NewIndexOutOfRangeException("index must be non-negative and less than or equal to the length.")
    If count < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum), "Count", count)
        
    Dim l As Long
    l = Len(value)
    If l > 0 And count > 0 Then
        mLength = InsertSpace(index, l * count)
        l = l * 2
        index = index * 2
        Do While count > 0
            Call CopyMemory(ByVal mStringPtr + index, ByVal StrPtr(value), l)
            index = index + l
            count = count - 1
        Loop
    End If
    Set Insert = Me
End Function

''
' Inserts an array of characters into the string value.
'
' @param index The index in the string value to start inserting.
' @param Chars The characters to be inserted.
' @param startindex The index in the Characters to start inserting from.
' @param count The number of characters to be inserted.
' @return This instance of StringBuilder.
'
Public Function InsertChars(ByVal index As Long, ByRef Chars() As Integer, Optional ByRef startIndex As Variant, Optional ByRef count As Variant) As StringBuilder
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    
    result = GetOptionalArrayRange(SAPtr(Chars), startIndex, ElemIndex, count, ElemCount)
    If result <> NO_ERROR Then _
        ThrowArrayRangeException result, "Chars", ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex)
    If index < 0 Or index > mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Index must be between 0 and then length of the current instance.", "Index", index)
    
    mLength = InsertSpace(index, ElemCount)
    Call CopyMemory(mString(index), Chars(ElemIndex), ElemCount * 2)
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
        Throw Cor.NewArgumentOutOfRangeException("Length cannot be set less than 0.", "Length", Length)
    
    If RHS > mLength Then
        If RHS > mCapacity Then Call EnsureCapacity(RHS)
        Call Fill(mLength, mFillChar, mCapacity - mLength)
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
Public Function Remove(ByVal startIndex As Long, ByVal count As Long) As StringBuilder
    Dim result As Long
    result = VerifyListRange(mLength, startIndex, count)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, startIndex, "StartIndex", count, "Count")
    
    If count > 0 Then
        Call CopyMemory(ByVal mStringPtr + startIndex * 2, ByVal mStringPtr + (startIndex + count) * 2, (mLength - (startIndex + count)) * 2)
        mLength = mLength - count
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
Public Function Replace(ByVal OldValue As String, ByVal NewValue As String, Optional ByRef startIndex As Variant, Optional ByRef count As Variant) As StringBuilder
    OldStrSA.pvData = StrPtr(OldValue)
    OldStrSA.cElements = Len(OldValue)

    If OldStrSA.cElements = 0 Then
        Set Replace = Me
        Exit Function
    End If

    Dim MatchCount      As Long
    Dim MatchIndexes()  As Long
    MatchCount = FindMatches(MatchIndexes, startIndex, count)

    ' If we have matches then we need to replace them.
    If MatchCount > 0 Then
        Dim NewLength As Long
        NewLength = mLength - ((Len(OldValue) - Len(NewValue)) * MatchCount)

        Select Case NewLength
            Case mLength:       Call ReplaceEqualLength(NewValue, MatchIndexes, MatchCount)
            Case Is < mLength:  Call ReplaceSmallerLength(NewValue, MatchIndexes, MatchCount, Len(OldValue))
            Case Else:          Call ReplaceLargerLength(NewValue, MatchIndexes, MatchCount, Len(OldValue), NewLength)
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
Public Function ToString(Optional ByRef startIndex As Variant, Optional ByRef Length As Variant) As String
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalListRange(mLength, startIndex, ElemIndex, Length, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Length", IsMissing(startIndex))
    
    If mLength = 0 Then
        ToString = ""
    Else
        ToString = SysAllocStringLen(mStringPtr + ElemIndex * 2, ElemCount)
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef s As String, ByRef startIndex As Variant, ByRef count As Variant, ByVal Capacity As Long)
    Call CoTaskMemFree(mStringPtr)
    
    Dim ElemCount As Long
    ElemCount = GetOptionalLong(count, Len(s))
    
    Dim ElemIndex As Long
    ElemIndex = GetOptionalLong(startIndex, 0)
    
    Dim result As Long
    result = VerifyListRange(Len(s), ElemIndex, ElemCount)
    If result <> NO_ERROR Then Call ThrowListRangeException(result, ElemIndex, "StartIndex", ElemCount, "Count", IsMissing(startIndex))
    
    If ElemCount > Capacity Then Capacity = ElemCount
    If Capacity < DEF_CAPACITY Then Capacity = DEF_CAPACITY
    
    mStringPtr = CoTaskMemAlloc(Capacity * 2)
    mCapacity = Capacity
    mLength = ElemCount
    If ElemCount > 0 Then Call CopyMemory(ByVal mStringPtr, ByVal StrPtr(s) + ElemIndex * 2, ElemCount * 2)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mFillChar = DEF_FILLCHAR
    mStringPtr = CoTaskMemAlloc(DEF_CAPACITY * 2)
    If mStringPtr = vbNullPtr Then Throw New OutOfMemoryException
    mCapacity = DEF_CAPACITY
    
    With OldStrSA
        .cbElements = 2
        .cDims = 1
        .cElements = &H7FFFFFFF
    End With
    SAPtr(OldStr) = VarPtr(OldStrSA)
    BufStrSA = OldStrSA
    SAPtr(BufStr) = VarPtr(BufStrSA)
    mStringSA = OldStrSA
    SAPtr(mString) = VarPtr(mStringSA)
    mStringSA.pvData = mStringPtr
    
    Call InitWordBuffer(mChars, 0, &H7FFFFFFF)
End Sub

Private Sub Class_Terminate()
    Call CoTaskMemFree(mStringPtr)
    SAPtr(OldStr) = 0
    SAPtr(BufStr) = 0
    SAPtr(mString) = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(value As Variant) As Boolean
    IObject_Equals = Equals(value)
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
Private Sub Fill(ByVal index As Long, ByVal Char As Integer, ByVal count As Long)
    Dim i As Long
    
    For i = index To index + count - 1
        mString(i) = Char
    Next i
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException("Invalid format specified.")
End Sub

Private Function InsertSpace(ByVal index As Long, ByVal size As Long) As Long
    InsertSpace = mLength + size
    If InsertSpace > mCapacity Then Call EnsureCapacity(InsertSpace)
    If index < mLength Then
        ' Create space in the string. The new space is uninitialized.
        Call CopyMemory(ByVal mStringPtr + (index + size) * 2, ByVal mStringPtr + index * 2, (mLength - index) * 2)
    End If
End Function


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
Friend Sub InternalAppendFormat(ByVal Provider As IFormatProvider, ByRef Format As String, ByRef args() As Variant)
    Dim CustomFormatter     As ICustomFormatter
    Dim HasCustomFormatter  As Boolean
    Dim index               As Long
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
    Do While index < Max
        Select Case mChars.Data(index)
            Case OPEN_BRACE
                ' potential escape or format info.
                
                If index > SubStringStart Then
                    ' If we have some characters, just append them now
                    ' even if the "{" may be escaped. Most times it won't be.
                    Call AppendChars(mChars.Data, SubStringStart, index - SubStringStart)
                End If
                
                index = index + 1
                
                ' We test for escape by checking the next character for a "{".
                ' But if we're out of characters, then we ended with a "{",
                ' and that is an invalid format.
                If index = Max Then Call FormatError
                
                ' We have enough characters to test for possible escape.
                If mChars.Data(index) = OPEN_BRACE Then
                    ' escape it by setting the start of the substring to the second one.
                    SubStringStart = index
                Else
                    Dim ArgIndex        As Long
                    Dim ArgWidth        As Long
                    Dim ArgFormatting   As String
                    
                    ' must be format info within a "{}", so call the parser for it.
                    index = ParseArgumentFormatting(index, ArgIndex, ArgWidth, ArgFormatting)
                    If ArgIndex < 0 Or ArgIndex > UBound(args) Then _
                        Throw Cor.NewFormatException("The argument index must be from 0 to " & UBound(args) & ".")
                        
                    ' append formatted argument here.
                    Dim s As String
                    
                    ' Set this because we are in a loop and always
                    ' need to begin with a vbNullString.
                    s = vbNullString
                    
                    ' Give a custom formatter first shot at formatting.
                    If HasCustomFormatter Then
                        s = CustomFormatter.Format(ArgFormatting, args(ArgIndex), Provider)
                    End If
                    
                    ' We distinguish between a null string and an empty string.
                    ' Null string indicates that the value was not formatted.
                    If StrPtr(s) = vbNullPtr Then
                        s = Convert.ToString(args(ArgIndex), ArgFormatting, Provider)
                    End If
                    
                    ' Add leading padding.
                    If ArgWidth > Len(s) Then
                        Call AppendChar(vbSpace, ArgWidth - Len(s))
                    End If
                    
                    ' Finally append the formatted value to the string.
                    Call AppendQuick(s)
                    
                    ' Add trailing padding.
                    If -ArgWidth > Len(s) Then
                        Call AppendChar(vbSpace, (-ArgWidth) - Len(s))
                    End If
                    
                    SubStringStart = index + 1
                    
                End If
            Case CLOSE_BRACE
                ' must be escaped or else it is a format error.
                
                index = index + 1
                
                ' Need enough characters to test for escape.
                If index = Max Then Call FormatError
                
                If mChars.Data(index) = CLOSE_BRACE Then
                    ' escape it and add the substring to this point.
                    ' The substring ends one character back because we
                    ' don't want to include both "}}" in the substring.
                    If index - 1 > SubStringStart Then
                        ' Append the current substring.
                        Call AppendChars(mChars.Data, SubStringStart, index - SubStringStart - 1)
                    End If
                    SubStringStart = index
                Else
                    ' we failed to escape it, so there was no matching "{".
                    Call FormatError
                End If
        End Select
        index = index + 1
    Loop
    
    If index > SubStringStart Then
        ' If we have some characters, just append them now.
        Call AppendChars(mChars.Data, SubStringStart, index - SubStringStart)
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
Private Function ParseArgumentFormatting(ByVal index As Long, ByRef ArgIndex As Long, ByRef ArgWidth As Long, ByRef ArgFormat As String) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    ' Get the argument index.
    ArgIndex = ParseNumber(index)
    
    index = ConsumeSpaces(index)
    
    With mChars
        ' Check if a column size is being defined.
        If .Data(index) = vbComma Then
            index = ConsumeSpaces(index + 1)
            
            ' Get the column width and aligment. Remember
            ' negative values mean align left in the column.
            ArgWidth = ParseNumber(index)
            
            index = ConsumeSpaces(index)
        Else
            ' Make sure they know there was no width specified.
            ArgWidth = 0
        End If
        
        ' Check if formatting is being specified.
        If .Data(index) = vbColon Then
            Dim startIndex As Long
            
            ' everything from here to the "}" is part of the argument formatting.
            index = index + 1
            startIndex = index
            Do
                ' can never reach the end until after a "}".
                If index = Max Then FormatError
                If .Data(index) = CLOSE_BRACE Then
                    ArgFormat = SysAllocStringLen(VarPtr(.Data(startIndex)), index - startIndex)
                    Exit Do
                End If
                index = index + 1
            Loop
        Else
            ' Make sure we let them know that no formatting was supplied.
            ArgFormat = vbNullString
        End If
        
        ' if we get here we should have reached the closing "}".
        If .Data(index) <> CLOSE_BRACE Then FormatError
    End With
    ParseArgumentFormatting = index
End Function

''
' Advances the index to the next non-space character.
'
' @param The index to the first space to be consumed.
' @return The index of the first character after the contiguous
' stream of spaces have been consumed.
'
Private Function ConsumeSpaces(ByVal index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Do
        ' Can never hit the end consuming spaces.
        If index = Max Then Call FormatError
        If mChars.Data(index) <> vbSpace Then Exit Do
        index = index + 1
    Loop
    
    ConsumeSpaces = index
End Function

''
' Parses out a positive or negative number.
'
' @param Index The index to the first digit or the "-" preceeding the first digit.
' @return The number that was parsed.
' @remarks The Index is passed by Reference and will be altered.
'
Private Function ParseNumber(ByRef index As Long) As Long
    Dim Max As Long
    Max = mChars.SA.cElements
    
    Dim Sign As Long
    Sign = 1
    
    Dim FirstCharIndex As Long
    FirstCharIndex = index
    
    Dim Number As Long
    Do
        ' We can never hit the end parsing a number.
        If index = Max Then Call FormatError
        
        Dim ch As Integer
        ch = mChars.Data(index)
        Select Case ch
            Case vbZero To vbNine
                ' Keep adding digits to the number
                Number = Number * 10 + ch - vbZero
            Case vbMinus
                ' The "-" can only be the first character.
                If index <> FirstCharIndex Then Call FormatError
                Sign = -1
            Case Else
                ' we have reached the end of the number, so exit.
                Exit Do
        End Select
        index = index + 1
    Loop
    
    ParseNumber = Number * Sign
End Function

''
' Loops through the string buffer finding all matching occurrences of the old value
' and storing the index for each match found.
'
' Returns the number of matches found.
Private Function FindMatches(ByRef MatchIndexes() As Long, ByRef startIndex As Variant, ByRef count As Variant) As Long
    Dim ElemCount   As Long
    Dim ElemIndex   As Long
    Dim result      As Long
    result = GetOptionalLongPair(startIndex, 0, ElemIndex, count, mLength, ElemCount)
    If result <> NO_ERROR Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(result))
    
    result = VerifyListRange(mLength, ElemIndex, ElemCount)
    If result <> NO_ERROR Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(result))

    ' We'll cache this since we'll be hitting it constantly.
    Dim FirstOldValueChar As Long
    FirstOldValueChar = OldStr(0)
    
    ' Don't need to search at the end if the old value
    ' can't fit. If the end matches the old value, then
    ' the inner loop will finish out the matching.
    ElemCount = ElemCount - OldStrSA.cElements + 1
    
    Dim MatchCount As Long
    ReDim MatchIndexes(0 To mLength)
    
    ' Loop through the current StringBuilder string trying to find
    ' all the matches against the old value to be replaced. If a
    ' match is found, then the index to that match is stored for
    ' future usage when replacing the matched old values with the new.
    Do While ElemCount > 0
        ' Found the start of a possible match.
        If mString(ElemIndex) = FirstOldValueChar Then
            ' We'll be opitmistic and assume a match.
            Dim Matched As Boolean
            Matched = True
            
            ' We optimistically save the index now
            ' because we'll be changing it while
            ' matching occurs.
            MatchIndexes(MatchCount) = ElemIndex
            
            ' Start with the second character.
            Dim i As Long
            For i = 1 To OldStrSA.cElements - 1
                If mString(ElemIndex + i) <> OldStr(i) Then
                    Matched = False
                    Exit For
                End If
            Next i
                
            If Matched Then
                MatchCount = MatchCount + 1
                ElemIndex = ElemIndex + OldStrSA.cElements - 1
                ElemCount = ElemCount - OldStrSA.cElements + 1
            End If
        End If
        
        ElemIndex = ElemIndex + 1
        ElemCount = ElemCount - 1
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
        Call CopyMemory(mString(MatchIndexes(i)), ByVal NewValuePtr, AmountToCopy)
    Next i
End Sub

''
' This will simply use the original string buffer as the source and destination buffers.
'
' @remarks Since the resulting string will be smaller, there is no fear of overwritting
' any original text that needs to end up in the final string.
'
Private Sub ReplaceSmallerLength(ByRef NewValue As String, ByRef MatchIndexes() As Long, ByVal MatchCount As Long, ByVal OldValueLength As Long)
    Call ReplaceMatches(mStringPtr, mStringPtr, NewValue, MatchIndexes, MatchCount, OldValueLength)
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
    NewBuffer = CoTaskMemAlloc(NewCapacity * 2)
    If NewBuffer = vbNullPtr Then Throw New OutOfMemoryException
    
    Call ReplaceMatches(mStringPtr, NewBuffer, NewValue, MatchIndexes, MatchCount, OldValueLength)
    
    Call CoTaskMemFree(mStringPtr)
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
            Call CopyMemory(ByVal lpDestBuffer, ByVal lpSourceBuffer + (OriginalIndex * 2), AmountToCopy * 2)
            lpDestBuffer = lpDestBuffer + AmountToCopy * 2
        End If
        
        If NewValueLen > 0 Then
            Call CopyMemory(ByVal lpDestBuffer, ByVal NewValuePtr, NewValueLen)
            lpDestBuffer = lpDestBuffer + NewValueLen
        End If
        
        OriginalIndex = OriginalIndex + AmountToCopy + OldValueLength
    Next i
    
    If OriginalIndex < mLength Then
        Call CopyMemory(ByVal lpDestBuffer, ByVal lpSourceBuffer + (OriginalIndex * 2), (mLength - OriginalIndex) * 2)
    End If
End Sub

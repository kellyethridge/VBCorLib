VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'    Module: StringWriter
'

''
' Provides a set of functions for writing to a string.
'
' @see Constructors
' @see TextWriter
'
Option Explicit
Implements IObject
Implements TextWriter

Private mSB         As New StringBuilder
Private mProvider   As IFormatProvider
Private mEncoding   As New UnicodeEncoding
Private mNewLine    As String
Private mIsClosed   As Boolean



''
' Returns the encoding in which the output is written.
'
' @return The encoding of the output string.
'
Public Property Get Encoding() As Encoding
    Set Encoding = mEncoding
End Property

''
' Returns the format provider the writer uses for formatting values.
'
' @return The format provider used by the writer.
' @remarks The format provider used is the returned by the current culture.
'
Public Property Get FormatProvider() As IFormatProvider
    If mProvider Is Nothing Then
        Set mProvider = CultureInfo.CurrentCulture
    End If
    
    Set FormatProvider = mProvider
End Property

''
' Returns the string currently used to represent a new-line set of characters.
'
' @return The new-line set of characters.
' @remarks The default is CarriageReturn-LineFeed.
'
Public Property Get NewLine() As String
    NewLine = mNewLine
End Property

''
' Sets the string to be used as the new-line set of characters.
'
' @param Value The new-line set of characters.
' @remarks The default is CarriageReturn-LineFeed.
'
Public Property Let NewLine(ByRef Value As String)
    mNewLine = Value
End Property

''
' Closes the writer
'
Public Sub CloseWriter()
    mIsClosed = True
End Sub

''
' Writes any buffered data to the underlying string.
'
' @remarks Since the data is directly written to the underlying string,
' this function does nothing. It is here to maintain interface consistency.
'
Public Sub Flush()
    ' does nothing
End Sub

''
' Returns the underlying stringbuilder being used to write to the string.
'
' @return The StringBuilder being used to write to the string.
'
Public Function GetStringBuilder() As StringBuilder
    Set GetStringBuilder = mSB
End Function

''
' Writes a value to the string.
'
' @param Value The value to be written to the stream.
' @param args The arguments to be used in formatting the Value.
' @remarks The <i>args</i> are used when argument indexes are supplied
' in the value to be written. Indexes are indicate by enclosing the
' specific index in curly braces {}. The indexes start at 0.
'
Public Sub WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    WriteCore Value, SwappedArgs
End Sub

Public Sub WriteChars(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    CheckValidSingleDimArray Chars, Parameter_Chars
    
    Dim Range As ListRange
    Range = OptionalRange(Index, Count, LBound(Chars), CorArray.LengthFirstDim(Chars))
    
    If Range.Index < LBound(Chars) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.NegativeNumber "Count"
    If Range.Index + Range.Count - 1 > UBound(Chars) Then _
        Error.InvalidOffsetLength
    
    mSB.Append Chars, Index, Count
End Sub

''
' Writes a value to the string, followed by a new-line set of characters.
'
' @param Values The value to be written to the string
' @remarks The first argument in <i>values</i> is the output to the
' string. Any additional arguments are used as formatting information.
' An argument can be refered to through formatting indexes. A
' formatting index is an integer enclosed in curly braces {}. The
' indexes start at 0.
' <p>Numbers and Dates will be formatted to the culture that was
' passed in using Cor.NewStringWriter. If no culture was specified, then
' the current culture for this computer is used.</p>
'
Public Sub WriteLine(ParamArray Values() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(12)
    WriteLineCore SwappedArgs
End Sub

Public Sub WriteCharsLine(ByRef Chars() As Integer, Optional ByRef Index As Variant, Optional ByRef Count As Variant)
    WriteChars Chars, Index, Count
    mSB.AppendString mNewLine
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = mSB.ToString
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param Value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mNewLine = Environment.NewLine
End Sub

Friend Sub Init(ByVal sb As StringBuilder, ByVal Provider As IFormatProvider)
    Set mSB = sb
    Set mProvider = Provider
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyIsOpen()
    If mIsClosed Then
        Error.ObjectDisposed ObjectDisposed_WriterClosed
    End If
End Sub

Private Sub WriteCore(ByRef Value As Variant, ByRef Args() As Variant)
    VerifyIsOpen
    
    Dim s As String
    s = Object.ToString(Value, Provider:=FormatProvider)
    
    If SAPtr(Args) = vbNullPtr Then
        mSB.AppendString s
    ElseIf UBound(Args) < LBound(Args) Then
        mSB.AppendString s
    Else
        mSB.InternalAppendFormat FormatProvider, s, Args
    End If
End Sub

Private Sub WriteLineCore(ByRef Args() As Variant)
    VerifyIsOpen
    
    Dim ub  As Long
    ub = UBound(Args)
    
    If ub >= 0 Then
        Dim s As String
        s = Object.ToString(Args(0), Provider:=FormatProvider) & mNewLine
        
        If ub > 0 Then
            CopyMemory Args(0), Args(1), (ub - LBound(Args)) * SizeOfVariant
            ZeroMemory Args(ub), SizeOfVariant
            mSB.InternalAppendFormat FormatProvider, s, Args
        Else
            mSB.AppendString Object.ToString(Args(0), Provider:=FormatProvider) & mNewLine
        End If
    Else
        mSB.AppendString mNewLine
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
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
'   TextWriter
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TextWriter_CloseWriter()
    CloseWriter
End Sub

Private Property Get TextWriter_Encoding() As Encoding
    Set TextWriter_Encoding = Encoding
End Property

Private Function TextWriter_Equals(Value As Variant) As Boolean
    TextWriter_Equals = Equals(Value)
End Function

Private Sub TextWriter_Flush()
    Flush
End Sub

Private Property Get TextWriter_FormatProvider() As IFormatProvider
    Set TextWriter_FormatProvider = FormatProvider
End Property

Private Function TextWriter_GetHashCode() As Long
    TextWriter_GetHashCode = GetHashCode
End Function

Private Property Let TextWriter_NewLine(ByVal RHS As String)
    NewLine = RHS
End Property

Private Property Get TextWriter_NewLine() As String
    TextWriter_NewLine = NewLine
End Property

Private Function TextWriter_ToString() As String
    TextWriter_ToString = ToString
End Function

Private Sub TextWriter_WriteChars(Chars() As Integer, Optional Index As Variant, Optional Count As Variant)
    WriteChars Chars, Index, Count
End Sub

Private Sub TextWriter_WriteCharsLine(Chars() As Integer, Optional Index As Variant, Optional Count As Variant)
    WriteCharsLine Chars, Index, Count
End Sub

Private Sub TextWriter_WriteLine(ParamArray Value() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(12)
    WriteLineCore SwappedArgs
End Sub

Private Sub TextWriter_WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    Dim SwappedArgs() As Variant
    Helper.Swap4 ByVal ArrPtr(SwappedArgs), ByVal Helper.DerefEBP(16)
    WriteCore Value, SwappedArgs
End Sub

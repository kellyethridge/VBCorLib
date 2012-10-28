VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Convert"
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
'    Module: Convert
'

''
' Provides methods used to encode and decode byte arrays to and from base-64 encoded characters.
'
' @remarks
' To access the methods of this class, simply call the method using the <b>Convert.*</b> syntax.
' <pre>
'''This example takes a byte array containing 10 values
'''from 1 to 10 and encodes it into a Base-64 string encoding.
'''The encoded string is then decoded back into the original
'''byte array and displayed for view.
'
''Private Sub Main()
''    Dim original() As Byte
''    Dim decoded() As Byte
''    Dim s As String
'
''    ' Create a byte array containing values 1 to 10.
''    original = Cor.NewBytes( 1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
'
''    ' Display the original byte array values
''    Console.WriteLine "Original Bytes"
''    DisplayBytes original
'
''    ' Encode the byte array into a Base-64 Encoded string.
''    s = Convert.ToBase64String(original)
'
''    ' Display Base-64 Encoded string.
''    Console.WriteLine
''    Console.WriteLine "Base-64 Encoded String"
''    Console.WriteLine " " & s
'
''    ' Decode Base-64 string into byte array.
''    decoded = Convert.FromBase64String(s)
'
''    ' Display decoded byte array.
''    Console.WriteLine "Bytes After Decoding"
''    DisplayBytes decoded
'
''    ' Wait for user to press enter
''    Console.ReadLine
''End Sub
'
'''Displays the contents of a byte array.
''Private Sub DisplayBytes(ByRef bytes() As Byte)
''    Dim i As Long
'
''    For i = LBound(bytes) To UBound(bytes)
''        Console.WriteValue " " & bytes(i)
''    Next i
''End Sub
'
'''This code produces the following output.
'
'''original bytes
''' 1 2 3 4 5 6 7 8 9 10
'''Base-64 Encoded String
''' AQIDBAUGBwgJCg==
'''Bytes After Decoding
''' 1 2 3 4 5 6 7 8 9 10
' </pre>
'
Option Explicit


Private mChars As WordBuffer


''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character string consisting of base-64 characters. Parameters specify
' the subsets as offsets of the input array and the number of elements in the input array to process.
'
' @param Bytes The bytes to be converted to a base-64 character array.
' @param Index Index within <i>Bytes</i> to begin encoding to base-64 characters.
' @param Count The number of bytes to encode.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return A string containg the byte array encoded as base-64 characters.
'
Public Function ToBase64String(ByRef Bytes() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal InsertLineBreaks As Boolean = False) As String
    Dim ElemIndex   As Long
    Dim ElemCount   As Long
    Dim Result      As Long
    Result = GetOptionalArrayRange(SAPtr(Bytes), Index, ElemIndex, Count, ElemCount)
    If Result <> NO_ERROR Then ThrowArrayRangeException Result, "Bytes", ElemIndex, "Index", ElemCount, "Count", IsMissing(Index)
    
    Dim Size As Long
    Size = ElemCount + ElemCount \ 3
    If (Size Mod 4) <> 0 Then Size = Size + (4 - (Size Mod 4))
    
    If InsertLineBreaks Then Size = Size + ((Size - 1) \ 76) * 2
    
    ToBase64String = SysAllocStringLen(0, Size)
    mChars.SA.pvData = StrPtr(ToBase64String)
    mChars.SA.cElements = Len(ToBase64String)
    Call ToBase64CharArray(Bytes, ElemIndex, ElemCount, mChars.Data, 0, InsertLineBreaks)
End Function

''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character array consisting of base 64 digits. Parameters specify
' the subsets as offsets of the input and output arrays and the number of elements in the input array.
'
' @param InArray The bytes to be converted to a base-64 character array.
' @param OffsetIn Offset within <i>InArray</i> to begin encoding to base-64 characters.
' @param Length The number of bytes to encode.
' @param OutArray The output character array containing the base-64 characters.
' @param OffsetOut The starting index in <i>OutArray</i> to place base-64 characters.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return The number of base-64 characters created.
' @remarks Added .NET 2.0 ability to insert linebreaks every 76 characters for MIME formatting.
'
Public Function ToBase64CharArray(ByRef InArray() As Byte, ByVal OffsetIn As Long, ByVal Length As Long, ByRef OutArray() As Integer, ByVal OffsetOut As Long, Optional ByVal InsertLineBreaks As Boolean = False) As Long
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InArray), OffsetIn, Length)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InArray", OffsetIn, "OffsetIn", Length, "Length")
    
    Result = VerifyArrayRange(SAPtr(OutArray), OffsetOut, Length)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "OutArray", OffsetOut, "OffsetOut", Length, "Length")
    
    Dim Size As Long
    Size = OffsetIn + (Length - (Length Mod 3))
    
    Dim CharCapacity As Long
    CharCapacity = Length + Length \ 3
    
    If (CharCapacity Mod 4) <> 0 Then
        CharCapacity = CharCapacity + (4 - (CharCapacity Mod 4))
    End If
    
    If InsertLineBreaks Then CharCapacity = CharCapacity + ((CharCapacity - 1) \ 76) * 2
    
    If OffsetOut + CharCapacity - 1 > UBound(OutArray) Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), "OutArray")
    
    
    Dim i           As Long
    Dim CharCount   As Long
    For i = OffsetIn To Size - 1 Step 3
        If InsertLineBreaks And (CharCount = 76) Then
            OutArray(OffsetOut) = vbReturn
            OutArray(OffsetOut + 1) = vbLineFeed
            OffsetOut = OffsetOut + 2
            CharCount = 0
        End If
        
        OutArray(OffsetOut) = Base64Bytes((InArray(i) And &HFC) \ 4)
        OutArray(OffsetOut + 1) = Base64Bytes(((InArray(i) And &H3) * &H10) Or ((InArray(i + 1) And &HF0) \ &H10))
        OutArray(OffsetOut + 2) = Base64Bytes(((InArray(i + 1) And &HF) * 4) Or ((InArray(i + 2) And &HC0) \ &H40))
        OutArray(OffsetOut + 3) = Base64Bytes(InArray(i + 2) And &H3F)
        OffsetOut = OffsetOut + 4
        If InsertLineBreaks Then CharCount = CharCount + 4
    Next i
    
    If InsertLineBreaks Then
        Select Case Length Mod 3
            Case 1, 2
                If CharCount = 76 Then
                    OutArray(OffsetOut) = vbReturn
                    OutArray(OffsetOut + 1) = vbLineFeed
                    OffsetOut = OffsetOut + 2
                End If
        End Select
    End If
    
    ' we need to pad the end of the string with "=" to fill out
    ' the current unfilled buffer.
    Select Case Length Mod 3
        Case 1
            OutArray(OffsetOut) = Base64Bytes((InArray(Size) And &HFC) \ 4)
            OutArray(OffsetOut + 1) = Base64Bytes((InArray(Size) And &H3) * &H10)
            OutArray(OffsetOut + 2) = vbEqual
            OutArray(OffsetOut + 3) = vbEqual
            OffsetOut = OffsetOut + 4
        Case 2
            OutArray(OffsetOut) = Base64Bytes((InArray(Size) And &HFC) \ 4)
            OutArray(OffsetOut + 1) = Base64Bytes(((InArray(Size) And &H3) * &H10) Or ((InArray(Size + 1) And &HF0) \ &H10))
            OutArray(OffsetOut + 2) = Base64Bytes((InArray(Size + 1) And &HF) * 4)
            OutArray(OffsetOut + 3) = vbEqual
            OffsetOut = OffsetOut + 4
    End Select
    ToBase64CharArray = OffsetOut
End Function

''
' Decodes a base-64 character array to a byte array. The Offset and length specify a subset of
' the character array to be decoded.
'
' @param InArray A base-64 character array to be decoded to a byte array.
' @param Offset The character within the character array to begin decoding.
' @param Length The number of characters to be decoded into the byte array.
' @return A byte array that has been decoded from a base-64 character array.
'
Public Function FromBase64CharArray(ByRef InArray() As Integer, ByVal Offset As Long, ByVal Length As Long) As Byte()
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InArray), Offset, Length)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InArray", Offset, "Offset", Length, "Length")
    
    Dim ret() As Byte
    ret = cArray.CreateInstance(ciByte, Length) ' use length because 1 char equals 6 bits, so we will have extra room.

    Dim i           As Long
    Dim j           As Long
    Dim Bits        As Long
    Dim BitCount    As Long
    Dim ch          As Integer
    Dim CharCount   As Long
    Dim TermCount   As Long
    For i = Offset To Offset + Length - 1
        ch = InArray(i)
        Select Case ch
            Case vbEqual
                CharCount = CharCount + 1
                TermCount = TermCount + 1
                
            Case &H20, &H9, &HA, &HB, &HC, &HD, &H85, &HA0
                If TermCount > 0 Then _
                    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
                    
            Case vbLowerA To vbLowerZ, vbUpperA To vbUpperZ, vbZero To vbNine, vbPlus, vbForwardSlash
                If TermCount > 0 Then _
                    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
                
                Bits = (Bits * &H40) Or Base64CharToBits(InArray(i))
                BitCount = BitCount + 6
                If BitCount >= 8 Then
                    ret(j) = Helper.ShiftRight(Bits, BitCount - 8) And &HFF
                    j = j + 1
                    BitCount = BitCount - 8
                    Bits = Bits And (Powers(BitCount) - 1)
                End If
                CharCount = CharCount + 1
            
            Case Else
                Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
        End Select
        If TermCount > 2 Then _
            Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
    Next i
    
    If (CharCount Mod 4) <> 0 Then _
        Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidNumberOfCharacters))
    
    Do While BitCount >= 8
        ret(j) = Helper.ShiftRight(Bits, BitCount - 8) And &HFF
        j = j + 1
        BitCount = BitCount - 8
        Bits = Bits And (Powers(BitCount) - 1)
    Loop
    
    If j - 1 <> UBound(ret) Then ReDim Preserve ret(0 To j - 1)
    
    FromBase64CharArray = ret
End Function

''
' A base-64 string containing characters to be decoded to a byte array.
'
' @param s The base-64 string to be decoded.
' @return A byte array containing the decoded base-64 characters.
'
Public Function FromBase64String(ByRef s As String) As Byte()
    mChars.SA.pvData = StrPtr(s)
    mChars.SA.cElements = Len(s)
    FromBase64String = FromBase64CharArray(mChars.Data, 0, Len(s))
End Function

''
' Converts a datatype value to a string representation using any
' supplied formatting or provider arguments.
'
' @param Value The value to convert to a string.
' @param Format Formatting information for converting the value.
' @param Provider A formatting provider to help custom formatting.
' @return A string representation of the value.
'
Public Function ToString(ByRef Value As Variant, Optional ByVal Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim vt As Long
    vt = VarType(Value)
    If vt = vbVariant Then
        vt = MemLong(MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)) And &HFF
    End If
    
    Select Case vt
        Case vbLong, vbInteger, vbByte, vbDouble, vbSingle
            Dim nf As NumberFormatInfo
            If Not Provider Is Nothing Then Set nf = Provider.GetFormat("numberformatinfo")
            If nf Is Nothing Then Set nf = NumberFormatInfo.CurrentInfo
            ToString = nf.Format(Value, Format)
        Case vbDate
            Dim df As DateTimeFormatInfo
            If Not Provider Is Nothing Then Set df = Provider.GetFormat("datetimeformatinfo")
            If df Is Nothing Then Set df = DateTimeFormatInfo.CurrentInfo
            ToString = df.Format(Value, Format)
        Case vbObject
            If Value Is Nothing Then
                ToString = "Nothing"
            ElseIf TypeOf Value Is IFormattable Then
                Dim f As IFormattable
                Set f = Value
                ToString = f.ToString(Format, Provider)
            ElseIf TypeOf Value Is IObject Then
                Dim o As IObject
                Set o = Value
                ToString = o.ToString
            Else
                ToString = TypeName(Value)
            End If
        Case vbEmpty
            ToString = "Empty"
        Case vbNull
            ToString = "Null"
        Case vbError               ' Value is a missing optional value most likely.
            Exit Function
        Case Else
            ToString = Value
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call InitWordBuffer(mChars, 0, &H7FFFFFFF)
End Sub

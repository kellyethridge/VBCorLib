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
' Module: Convert
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

''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character string consisting of base-64 characters. Parameters specify
' the subsets as offsets of the input array and the number of elements in the input array to process.
'
' @param InArray The bytes to be converted to a base-64 character array.
' @param Offset Offset within <em>Bytes</em> to begin encoding to base-64 characters.
' @param Length The number of bytes to encode.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return A string containg the byte array encoded as base-64 characters.
'
Public Function ToBase64String(ByRef InArray() As Byte, Optional ByRef Offset As Variant, Optional ByRef Length As Variant, Optional ByVal InsertLineBreaks As Boolean = False) As String
    Dim Range   As ListRange
    Dim Chars() As Integer
    
    ValidateByteArray InArray, Parameter_InArray
    Range = OptionalStepRange(Offset, Length, LBound(InArray), Len1D(InArray))
    If Range.Index < LBound(InArray) Then _
        Error.ArgumentOutOfRange "Index", ArgumentOutOfRange_LBound
    If Range.Count < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    If Range.Index + Range.Count - 1 > UBound(InArray) Then _
        Error.ArgumentOutOfRange "", ArgumentOutOfRange_IndexLength
    
    ToBase64String = SysAllocStringLen(vbNullPtr, CalculateOutputLength(Range.Count, InsertLineBreaks))
    
    Chars = AllocChars(ToBase64String)
    ToBase64CharArrayCore InArray, Range.Index, Range.Count, Chars, 0, InsertLineBreaks
    FreeChars Chars
End Function

''
' Converts the value of a subset of a byte array to an equivalent
' subset of a Unicode character array consisting of base 64 digits. Parameters specify
' the subsets as offsets of the input and output arrays and the number of elements in the input array.
'
' @param InArray The bytes to be converted to a base-64 character array.
' @param IndexIn Offset within <i>InArray</i> to begin encoding to base-64 characters.
' @param Length The number of bytes to encode.
' @param OutArray The output character array containing the base-64 characters.
' @param IndexOut The starting index in <i>OutArray</i> to place base-64 characters.
' @param InsertLineBreaks Indicates if a vbCrLf should be inserted every 76 characters to
' conform with MIME line formatting.
' @return The number of base-64 characters created.
' @remarks Added .NET 2.0 ability to insert linebreaks every 76 characters for MIME formatting.
'
Public Function ToBase64CharArray(ByRef InArray() As Byte, ByVal IndexIn As Long, ByVal Length As Long, ByRef OutArray() As Integer, ByVal IndexOut As Long, Optional ByVal InsertLineBreaks As Boolean = False) As Long
    ValidateByteArray InArray, Parameter_InArray
    ValidateCharArray OutArray, Parameter_OutArray
    If IndexIn < LBound(InArray) Then _
        Error.ArgumentOutOfRange "IndexIn", ArgumentOutOfRange_LBound
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    If IndexIn + Length - 1 > UBound(InArray) Then _
        Error.ArgumentOutOfRange "IndexIn", ArgumentOutOfRange_IndexLength
    If IndexOut < LBound(OutArray) Then _
        Error.ArgumentOutOfRange "IndexOut", ArgumentOutOfRange_LBound
        
    Dim Size As Long
    Size = CalculateOutputLength(Length, InsertLineBreaks)
    
    If IndexOut + Size - 1 > UBound(OutArray) Then _
        Error.ArgumentOutOfRange "IndexOut", ArgumentOutOfRange_OffsetOut
    
    ToBase64CharArray = ToBase64CharArrayCore(InArray, IndexIn, Length, OutArray, IndexOut, InsertLineBreaks)
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
    ValidateCharArray InArray, Parameter_InArray
    If Offset < LBound(InArray) Then _
        Error.ArgumentOutOfRange "Offset", ArgumentOutOfRange_LBound
    If Length < 0 Then _
        Error.ArgumentOutOfRange "Length", ArgumentOutOfRange_NegativeLength
    If Offset + Length - 1 > UBound(InArray) Then _
        Error.ArgumentOutOfRange "Offset", ArgumentOutOfRange_IndexLength
    
    Dim Ret() As Byte
    SAPtr(Ret) = SafeArrayCreateVector(vbByte, 0, Length)

    Dim i           As Long
    Dim j           As Long
    Dim Bits        As Long
    Dim BitCount    As Long
    Dim Ch          As Integer
    Dim CharCount   As Long
    Dim TermCount   As Long
    For i = Offset To Offset + Length - 1
        Ch = InArray(i)
        Select Case Ch
            Case vbEqualChar
                CharCount = CharCount + 1
                TermCount = TermCount + 1
                
            Case &H20, &H9, &HA, &HB, &HC, &HD, &H85, &HA0
                If TermCount > 0 Then _
                    Error.Format Format_InvalidBase64Character
                    
            Case vbLowerAChar To vbLowerZChar, vbUpperAChar To vbUpperZChar, vbZeroChar To vbNineChar, vbPlusChar, vbForwardSlashChar
                If TermCount > 0 Then _
                    Error.Format Format_InvalidBase64Character
                
                Bits = (Bits * &H40) Or Base64CharToBits(InArray(i))
                BitCount = BitCount + 6
                If BitCount >= 8 Then
                    Ret(j) = Helper.ShiftRight(Bits, BitCount - 8) And &HFF
                    j = j + 1
                    BitCount = BitCount - 8
                    Bits = Bits And (Powers(BitCount) - 1)
                End If
                CharCount = CharCount + 1
            
            Case Else
                Error.Format Format_InvalidBase64Character

        End Select
        If TermCount > 2 Then _
            Error.Format Format_InvalidBase64Character
    Next i
    
    If (CharCount Mod 4) <> 0 Then _
        Error.Format Format_InvalidNumberOfCharacters
    
    Do While BitCount >= 8
        Ret(j) = Helper.ShiftRight(Bits, BitCount - 8) And &HFF
        j = j + 1
        BitCount = BitCount - 8
        Bits = Bits And (Powers(BitCount) - 1)
    Loop
    
    If j - 1 <> UBound(Ret) Then ReDim Preserve Ret(0 To j - 1)
    
    FromBase64CharArray = Ret
End Function

''
' A base-64 string containing characters to be decoded to a byte array.
'
' @param s The base-64 string to be decoded.
' @return A byte array containing the decoded base-64 characters.
'
Public Function FromBase64String(ByRef s As String) As Byte()
    Dim Chars() As Integer
    Chars = AllocChars(s)
    
    On Error GoTo Catch
    FromBase64String = FromBase64CharArray(Chars, 0, Len(s))
    
    GoSub Finally
    Exit Function
    
Catch:
    GoSub Finally
    Throw
Finally:
    FreeChars Chars
    Return
End Function

''
' Converts a datatype value to a string representation using any
' supplied formatting or provider arguments.
'
' @param Value The value to convert to a string.
' @param Provider A formatting provider to help custom formatting.
' @return A string representation of the value.
'
Public Function ToString(ByRef Value As Variant, Optional ByVal Provider As IFormatProvider) As String
    ToString = Object.ToString(Value, Provider:=Provider)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToBase64CharArrayCore(ByRef InArray() As Byte, ByVal IndexIn As Long, ByVal Length As Long, ByRef OutArray() As Integer, ByVal IndexOut As Long, ByVal InsertLineBreaks As Boolean) As Long
    Dim MaxIndex As Long
    MaxIndex = IndexIn + (Length - (Length Mod 3))
    
    Dim i           As Long
    Dim CharCount   As Long
    For i = IndexIn To MaxIndex - 1 Step 3
        If InsertLineBreaks And (CharCount = 76) Then
            OutArray(IndexOut) = vbReturnChar
            OutArray(IndexOut + 1) = vbLineFeedChar
            IndexOut = IndexOut + 2
            CharCount = 0
        End If
        
        OutArray(IndexOut) = Base64Bytes((InArray(i) And &HFC) \ 4)
        OutArray(IndexOut + 1) = Base64Bytes(((InArray(i) And &H3) * &H10) Or ((InArray(i + 1) And &HF0) \ &H10))
        OutArray(IndexOut + 2) = Base64Bytes(((InArray(i + 1) And &HF) * 4) Or ((InArray(i + 2) And &HC0) \ &H40))
        OutArray(IndexOut + 3) = Base64Bytes(InArray(i + 2) And &H3F)
        IndexOut = IndexOut + 4
        If InsertLineBreaks Then CharCount = CharCount + 4
    Next i
    
    If InsertLineBreaks Then
        Select Case Length Mod 3
            Case 1, 2
                If CharCount = 76 Then
                    OutArray(IndexOut) = vbReturnChar
                    OutArray(IndexOut + 1) = vbLineFeedChar
                    IndexOut = IndexOut + 2
                End If
        End Select
    End If
    
    ' we need to pad the end of the string with "=" to fill out
    ' the current unfilled buffer.
    Select Case Length Mod 3
        Case 1
            OutArray(IndexOut) = Base64Bytes((InArray(MaxIndex) And &HFC) \ 4)
            OutArray(IndexOut + 1) = Base64Bytes((InArray(MaxIndex) And &H3) * &H10)
            OutArray(IndexOut + 2) = vbEqualChar
            OutArray(IndexOut + 3) = vbEqualChar
            IndexOut = IndexOut + 4
        Case 2
            OutArray(IndexOut) = Base64Bytes((InArray(MaxIndex) And &HFC) \ 4)
            OutArray(IndexOut + 1) = Base64Bytes(((InArray(MaxIndex) And &H3) * &H10) Or ((InArray(MaxIndex + 1) And &HF0) \ &H10))
            OutArray(IndexOut + 2) = Base64Bytes((InArray(MaxIndex + 1) And &HF) * 4)
            OutArray(IndexOut + 3) = vbEqualChar
            IndexOut = IndexOut + 4
    End Select
    ToBase64CharArrayCore = IndexOut
End Function

Private Function CalculateOutputLength(ByVal Length As Long, ByVal InsertLineBreaks As Boolean) As Long
    Dim Size As Long
    Size = Length + Length \ 3
    If (Size Mod 4) <> 0 Then
        Size = Size + (4 - (Size Mod 4))
    End If
    
    If InsertLineBreaks Then
        Size = Size + ((Size - 1) \ 76) * 2
    End If
    
    CalculateOutputLength = Size
End Function



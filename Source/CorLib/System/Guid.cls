VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guid"
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
' Module: Guid
'

''
' Represents a globally unique identifier (GUID).
'
' @remarks A GUID is a 128-bit integer (16 bytes) that can be used across all
' computers and networks wherever a unique identifier is required. Such an identifier
' has a very low probability of being duplicated.
' <p>A new <b>Guid</b> can be created using the <a href="Constructors.NewGuid.html">NewGuid</a> method
' or one of the factory methods in <a href="GuidStatic.html">GuidStatic</a>.
'
' @see GuidStatic
' @see Constructors
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const PropData1         As String = "Data1"
Private Const PropData2         As String = "Data2"
Private Const PropData3         As String = "Data3"
Private Const PropData4         As String = "Data4"
Private Const PropIsReadOnly    As String = "IsReadOnly"

''
' FormatOptions
'
' @param None none
' @param Dashes dashes
' @param Closure closure
Private Enum FormatOptions
    None = 0
    Dashes = 1
    Closure = 2
End Enum

Private mGuid       As VBGUID
Private mGuidCopy   As VBGUID
Private mIsReadOnly As Boolean


''
' Returns a pointer to the internal GUID structure.
'
' @return A long containing the pointer to the structure.
' @remarks This is provided to allow for APIs to access the
' actual guid structure directly, since they expect to be
' accessing a structure, not an object.
' <p>Guid style APIs declare the guid parameter something like
' <pre>ByRef g As VBGUID</pre>
' In order to use the VBCorLib <b>Guid</b> then the API would
' need to be declared something like
' <pre>ByVal ptrG As Long</pre>
' and pass in the Handle to the guid object.</p>
' <p>Exposing the handle is very dangerous in that the object
' is not immutable.</p>
'
Public Property Get Handle() As Long
    If mIsReadOnly Then
        mGuidCopy = mGuid
        Handle = VarPtr(mGuidCopy)
    Else
        Handle = VarPtr(mGuid)
    End If
End Property

''
' Returns if the guid is read-only.
'
' @return Returns <b>True</b> if the guid is locked, otherwise <b>False</b>.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mIsReadOnly
End Property

''
' Returns the guid value has an array of bytes.
'
' @return An array of bytes representing the guid.
'
Public Function ToByteArray() As Byte()
    Dim Result(0 To vbSizeOfGuid - 1) As Byte
    
    CopyMemory Result(0), mGuid, vbSizeOfGuid
    ToByteArray = Result
End Function

''
' Compares this guid to another.
'
' @param Value The guid to compare to this guid.
' @return A Long
'
Public Function CompareTo(ByRef Value As Variant) As Long
    Dim Other As Guid
    
    On Error GoTo Catch
    Set Other = Value
    If Other Is Nothing Then
        CompareTo = 1
    Else
        CompareTo = -Other.InternalCompare(mGuid)
    End If
    
    Exit Function
    
Catch:
    Error.Argument Arg_MustBeGuid, "Value"
End Function

''
' Returns a string representation of guid value.
'
' @param Format The format of the guid value to be returned.
' @param Provider A format provider used to format the guid.
' @return String representing of the guid.
' @remarks There are 5 format types: "D","B","N","P","X".<br>
' <pre>
' "D" = 00000000-0000-0000-0000-000000000000
' "B" = {00000000-0000-0000-0000-000000000000}
' "P" = (00000000-0000-0000-0000-000000000000)
' "N" = 00000000000000000000000000000000
' "X" = {0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}}
' </pre>
' <b>D</b> is the default format.
'
Public Function ToString(Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim FormatCh As Integer
    
    Select Case Len(Format)
        Case 0
            FormatCh = vbUpperDChar
        Case 1
            FormatCh = AscW(Format)
        Case Else
            Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidGuidFormatSpecification))
    End Select
    
    Select Case FormatCh
        Case vbUpperDChar, vbLowerDChar
            ToString = ToStandardString(FormatOptions.Dashes)
        Case vbUpperBChar, vbLowerBChar
            ToString = ToStandardString(FormatOptions.Dashes Or FormatOptions.Closure, vbOpenBraceChar, vbCloseBraceChar)
        Case vbUpperPChar, vbLowerPChar
            ToString = ToStandardString(FormatOptions.Dashes Or FormatOptions.Closure, vbOpenParenChar, vbCloseParenChar)
        Case vbUpperNChar, vbLowerNChar
            ToString = ToStandardString(FormatOptions.None)
        Case vbUpperXChar, vbLowerXChar
            ToString = ToHexString
        Case Else
            Error.Format Format_InvalidGuidFormatSpecification
    End Select
End Function



''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As Guid
    
    On Error GoTo Catch
    Set Other = Value
    Equals = IsEqualGUID(mGuid, ByVal Other.Handle)

Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    With mGuid
        GetHashCode = .Data1 Xor AsLong(.Data2) Xor AsLong(.Data4(0)) Xor AsLong(.Data4(4))
    End With
End Function

Friend Function CreateReadOnly() As Guid
    Set CreateReadOnly = New Guid
    CreateReadOnly.InitReadOnly mGuid
End Function

Friend Function InternalCompare(ByRef g As VBGUID) As Long
    If Not EqualFields(mGuid.Data1, g.Data1, InternalCompare) Then Exit Function
    If Not EqualFields(mGuid.Data2, g.Data2, InternalCompare) Then Exit Function
    If Not EqualFields(mGuid.Data3, g.Data3, InternalCompare) Then Exit Function
    
    Dim i As Long
    For i = 0 To 7
        If Not EqualFields(mGuid.Data4(i), g.Data4(i), InternalCompare) Then Exit Function
    Next i
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant)
    Select Case VarType(Value)
        Case vbString
            Parse Value
        Case vbByteArray
            FromByteArray Value
        Case vbLong
            CopyMemory mGuid, ByVal CLng(Value), vbSizeOfGuid
        Case Else
            Error.Argument Format_GuidUnrecognized
    End Select
    
    mIsReadOnly = True
End Sub

Friend Sub InitReadOnly(ByRef g As VBGUID)
    mGuid = g
    mIsReadOnly = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FromByteArray(ByRef Value As Variant)
    ValidateArray Value, Parameter_Value
    
    If Len1D(Value) <> vbSizeOfGuid Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_GuidArrayCtor, vbSizeOfGuid))
    
    Dim Bytes() As Byte
    SAPtr(Bytes) = SAPtrV(Value)
    CopyMemory mGuid, Bytes(LBound(Bytes)), vbSizeOfGuid
    SAPtr(Bytes) = vbNullPtr
End Sub

Private Sub Parse(ByRef Value As Variant)
    If Not Guid.Parser.TryParse(CStr(Value), mGuid) Then
        Guid.Parser.ThrowError
    End If
End Sub

Private Function EqualFields(ByVal a As Long, ByVal b As Long, ByRef Result As Long) As Boolean
    Dim Upper8A As Long
    Dim Upper8B As Long
    
    Upper8A = ((a And &HFF000000) \ &H1000000) And &HFF
    Upper8B = ((b And &HFF000000) \ &H1000000) And &HFF
    Result = 0
    
    If Upper8A < Upper8B Then
        Result = -1
    ElseIf Upper8A > Upper8B Then
        Result = 1
    Else
        Dim Lower24A As Long
        Dim Lower24B As Long
        
        Lower24A = a And &HFFFFFF
        Lower24B = b And &HFFFFFF
        
        If Lower24A < Lower24B Then
            Result = -1
        ElseIf Lower24A > Lower24B Then
            Result = 1
        End If
    End If
    
    EqualFields = Result = 0
End Function

Private Function ToStandardString(ByVal Options As FormatOptions, Optional ByVal OpenChar As Integer, Optional ByVal CloseChar As Integer) As String
    Dim Size        As Long
    Dim Offset      As Long
    Dim Result      As String
    Dim Output()    As Integer
    
    Size = 32
    If Options And Dashes Then
        Size = Size + 4
    End If
    If Options And Closure Then
        Size = Size + 2
    End If
    
    Result = SysAllocStringLen(vbNullPtr, Size)
    Output = AllocChars(Result)
    
    If Options And Closure Then
        Output(0) = OpenChar
        Offset = 1
    End If
    
    Offset = HexToChars(Output, Offset, (mGuid.Data1 And &HFF000000) \ &H1000000, (mGuid.Data1 And &HFF0000) \ &H10000, False)
    Offset = HexToChars(Output, Offset, (mGuid.Data1 And &HFF00) \ &H100, mGuid.Data1, False)
    Offset = AddDash(Output, Offset, Options)
    Offset = HexToChars(Output, Offset, (mGuid.Data2 And &HFF00) \ &H100, mGuid.Data2, False)
    Offset = AddDash(Output, Offset, Options)
    Offset = HexToChars(Output, Offset, (mGuid.Data3 And &HFF00) \ &H100, mGuid.Data3, False)
    Offset = AddDash(Output, Offset, Options)
    Offset = HexToChars(Output, Offset, mGuid.Data4(0), mGuid.Data4(1), False)
    Offset = AddDash(Output, Offset, Options)
    Offset = HexToChars(Output, Offset, mGuid.Data4(2), mGuid.Data4(3), False)
    Offset = HexToChars(Output, Offset, mGuid.Data4(4), mGuid.Data4(5), False)
    Offset = HexToChars(Output, Offset, mGuid.Data4(6), mGuid.Data4(7), False)
    
    If Options And Closure Then
        Output(Offset) = CloseChar
    End If
    
    FreeChars Output
    ToStandardString = Result
End Function

Private Function AddDash(ByRef Output() As Integer, ByVal Offset As Long, ByVal Options As FormatOptions) As Long
    If Options And Dashes Then
        Output(Offset) = vbHyphenChar
        Offset = Offset + 1
    End If
    
    AddDash = Offset
End Function

Private Function ToHexString() As String
    Dim Result      As String
    Dim Output() As Integer
    
    Result = SysAllocStringLen(0, 68)
    Output = AllocChars(Result)
    
    Output(0) = vbOpenBraceChar
    Output(1) = vbZeroChar
    Output(2) = vbLowerXChar
    HexToChars Output, 3, (mGuid.Data1 And &HFF000000) \ &H1000000, (mGuid.Data1 And &HFF0000) \ &H10000, False
    HexToChars Output, 7, (mGuid.Data1 And &HFF00) \ &H100, mGuid.Data1, False
    Output(11) = vbCommaChar
    Output(12) = vbZeroChar
    Output(13) = vbLowerXChar
    HexToChars Output, 14, (mGuid.Data2 And &HFF00) \ &H100, mGuid.Data2, False
    Output(18) = vbCommaChar
    Output(19) = vbZeroChar
    Output(20) = vbLowerXChar
    HexToChars Output, 21, (mGuid.Data3 And &HFF00) \ &H100, mGuid.Data3, False
    Output(25) = vbCommaChar
    Output(26) = vbOpenBraceChar
    HexToChars Output, 27, mGuid.Data4(0), mGuid.Data4(1), True
    Output(36) = vbCommaChar
    HexToChars Output, 37, mGuid.Data4(2), mGuid.Data4(3), True
    Output(46) = vbCommaChar
    HexToChars Output, 47, mGuid.Data4(4), mGuid.Data4(5), True
    Output(56) = vbCommaChar
    HexToChars Output, 57, mGuid.Data4(6), mGuid.Data4(7), True
    Output(66) = vbCloseBraceChar
    Output(67) = vbCloseBraceChar
    
    FreeChars Output
    ToHexString = Result
End Function

Private Function HexToChars(ByRef Output() As Integer, ByVal Offset As Long, ByVal a As Long, ByVal b As Long, ByVal IsHex As Boolean) As Long
    a = a And &HFF
    b = b And &HFF
    
    If IsHex Then
        Output(Offset) = vbZeroChar
        Output(Offset + 1) = vbLowerXChar
        Offset = Offset + 2
    End If
    
    Output(Offset) = HexToChar(a \ &H10)
    Output(Offset + 1) = HexToChar(a)
    Offset = Offset + 2
    
    If IsHex Then
        Output(Offset) = vbCommaChar
        Output(Offset + 1) = vbZeroChar
        Output(Offset + 2) = vbLowerXChar
        Offset = Offset + 3
    End If
    
    Output(Offset) = HexToChar(b \ &H10)
    Output(Offset + 1) = HexToChar(b)
    
    HexToChars = Offset + 2
End Function

Private Function HexToChar(ByVal a As Long) As Integer
    a = a And &HF
    
    If a > 9 Then
        HexToChar = a - 10 + vbLowerAChar
    Else
        HexToChar = a + vbZeroChar
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mGuid.Data1 = .ReadProperty(PropData1, 0)
        mGuid.Data2 = .ReadProperty(PropData2, 0)
        mGuid.Data3 = .ReadProperty(PropData3, 0)
        
        Dim b() As Byte
        b = .ReadProperty(PropData4, 0)
        CopyMemory mGuid.Data4(0), b(0), 8
        
        mIsReadOnly = .ReadProperty(PropIsReadOnly, False)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropData1, mGuid.Data1
        .WriteProperty PropData2, mGuid.Data2
        .WriteProperty PropData3, mGuid.Data3
        .WriteProperty PropData4, mGuid.Data4
        .WriteProperty PropIsReadOnly, mIsReadOnly
    End With
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
'   IComparable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormattable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormattable_ToString(ByRef Format As String, ByVal Provider As IFormatProvider) As String
    IFormattable_ToString = ToString(Format, Provider)
End Function

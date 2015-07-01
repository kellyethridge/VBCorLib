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
'
' @see GuidStatic
'
Option Explicit
Implements IObject
Implements IComparable
Implements IFormattable

Private Const REGDB_E_CLASSNOTREG   As Long = &H80040154
Private Const REGDB_E_READREGDB     As Long = &H80040150

Private Enum FormatOptions
    None = 0
    Dashes = 1
    Closure = 2
End Enum

Private mGuid       As VBGUID
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
Attribute Handle.VB_MemberFlags = "40"
    If mIsReadOnly Then _
        Error.InvalidOperation InvalidOperation_ReadOnly
    
    Handle = VarPtr(mGuid)
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
    Dim Result(0 To corSizeOfGuid - 1) As Byte
    
    CopyMemory Result(0), mGuid, corSizeOfGuid
    ToByteArray = Result
End Function

''
' Compares this guid to another.
'
' @param Value The guid to compare to this guid.
' @return A Long
'
Public Function CompareTo(ByRef Value As Variant) As Long
    Dim other As Guid
    
    On Error GoTo Catch
    Set other = Value
    If other Is Nothing Then
        CompareTo = 1
    Else
        CompareTo = -other.InternalCompare(mGuid)
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
' <b>B</b> is the default format.
'
Public Function ToString(Optional ByRef Format As String, Optional ByVal Provider As IFormatProvider) As String
    Dim FormatCh As Integer
    
    Select Case Len(Format)
        Case 0
            FormatCh = corUpperDChar
        Case 1
            FormatCh = AscW(Format)
        Case Else
            Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidGuidFormatSpecification))
    End Select
    
    Select Case FormatCh
        Case corUpperDChar, corLowerDChar
            ToString = ToStandardString(FormatOptions.Dashes)
        Case corUpperBChar, corLowerBChar
            ToString = ToStandardString(FormatOptions.Dashes Or FormatOptions.Closure, corOpenBraceChar, corCloseBraceChar)
        Case corUpperPChar, corLowerPChar
            ToString = ToStandardString(FormatOptions.Dashes Or FormatOptions.Closure, corOpenParenChar, corCloseParenChar)
        Case corUpperNChar, corLowerNChar
            ToString = ToStandardString(FormatOptions.None)
        Case corUpperXChar, corLowerXChar
            ToString = ToHexString
        Case Else
            Error.Format Format_InvalidGuidFormatSpecification
    End Select
End Function

Private Function ToStandardString(ByVal Options As FormatOptions, Optional ByVal OpenChar As Integer, Optional ByVal CloseChar As Integer) As String
    Dim Size        As Long
    Dim Offset      As Long
    Dim Output      As String
    Dim GuidChars() As Integer
    
    Size = 32
    If Options And Dashes Then
        Size = Size + 4
    End If
    If Options And Closure Then
        Size = Size + 2
    End If
    
    Output = SysAllocStringLen(0, Size)
    GuidChars = AllocChars(Output)
    
    If Options And Closure Then
        GuidChars(0) = OpenChar
        Offset = 1
    End If
    
    Offset = HexToChars(GuidChars, Offset, (mGuid.Data1 And &HFF000000) \ &H1000000, (mGuid.Data1 And &HFF0000) \ &H10000, False)
    Offset = HexToChars(GuidChars, Offset, (mGuid.Data1 And &HFF00) \ &H100, mGuid.Data1, False)
    Offset = AddDash(GuidChars, Offset, Options)
    Offset = HexToChars(GuidChars, Offset, (mGuid.Data2 And &HFF00) \ &H100, mGuid.Data2, False)
    Offset = AddDash(GuidChars, Offset, Options)
    Offset = HexToChars(GuidChars, Offset, (mGuid.Data3 And &HFF00) \ &H100, mGuid.Data3, False)
    Offset = AddDash(GuidChars, Offset, Options)
    Offset = HexToChars(GuidChars, Offset, mGuid.Data4(0), mGuid.Data4(1), False)
    Offset = AddDash(GuidChars, Offset, Options)
    Offset = HexToChars(GuidChars, Offset, mGuid.Data4(2), mGuid.Data4(3), False)
    Offset = HexToChars(GuidChars, Offset, mGuid.Data4(4), mGuid.Data4(5), False)
    Offset = HexToChars(GuidChars, Offset, mGuid.Data4(6), mGuid.Data4(7), False)
    
    If Options And Closure Then
        GuidChars(Offset) = CloseChar
    End If
    
    FreeChars GuidChars
    ToStandardString = Output
End Function

Private Function AddDash(ByRef GuidChars() As Integer, ByVal Offset As Long, ByVal Options As FormatOptions) As Long
    If Options And Dashes Then
        GuidChars(Offset) = corHyphenChar
        Offset = Offset + 1
    End If
    
    AddDash = Offset
End Function

Private Function ToHexString() As String
    Dim Output      As String
    Dim GuidChars() As Integer
    
    Output = SysAllocStringLen(0, 68)
    GuidChars = AllocChars(Output)
    
    GuidChars(0) = corOpenBraceChar
    GuidChars(1) = corZeroChar
    GuidChars(2) = corLowerXChar
    HexToChars GuidChars, 3, (mGuid.Data1 And &HFF000000) \ &H1000000, (mGuid.Data1 And &HFF0000) \ &H10000, False
    HexToChars GuidChars, 7, (mGuid.Data1 And &HFF00) \ &H100, mGuid.Data1, False
    GuidChars(11) = corCommaChar
    GuidChars(12) = corZeroChar
    GuidChars(13) = corLowerXChar
    HexToChars GuidChars, 14, (mGuid.Data2 And &HFF00) \ &H100, mGuid.Data2, False
    GuidChars(18) = corCommaChar
    GuidChars(19) = corZeroChar
    GuidChars(20) = corLowerXChar
    HexToChars GuidChars, 21, (mGuid.Data3 And &HFF00) \ &H100, mGuid.Data3, False
    GuidChars(25) = corCommaChar
    GuidChars(26) = corOpenBraceChar
    HexToChars GuidChars, 27, mGuid.Data4(0), mGuid.Data4(1), True
    GuidChars(36) = corCommaChar
    HexToChars GuidChars, 37, mGuid.Data4(2), mGuid.Data4(3), True
    GuidChars(46) = corCommaChar
    HexToChars GuidChars, 47, mGuid.Data4(4), mGuid.Data4(5), True
    GuidChars(56) = corCommaChar
    HexToChars GuidChars, 57, mGuid.Data4(6), mGuid.Data4(7), True
    GuidChars(66) = corCloseBraceChar
    GuidChars(67) = corCloseBraceChar
    
    FreeChars GuidChars
    ToHexString = Output
End Function

Private Function HexToChars(ByRef GuidChars() As Integer, ByVal Offset As Long, ByVal a As Long, ByVal b As Long, ByVal IsHex As Boolean) As Long
    a = a And &HFF
    b = b And &HFF
    
    If IsHex Then
        GuidChars(Offset) = corZeroChar
        GuidChars(Offset + 1) = corLowerXChar
        Offset = Offset + 2
    End If
    
    GuidChars(Offset) = HexToChar(a \ &H10)
    GuidChars(Offset + 1) = HexToChar(a)
    Offset = Offset + 2
    
    If IsHex Then
        GuidChars(Offset) = corCommaChar
        GuidChars(Offset + 1) = corZeroChar
        GuidChars(Offset + 2) = corLowerXChar
        Offset = Offset + 3
    End If
    
    GuidChars(Offset) = HexToChar(b \ &H10)
    GuidChars(Offset + 1) = HexToChar(b)
    
    HexToChars = Offset + 2
End Function

Private Function HexToChar(ByVal a As Long) As Integer
    a = a And &HF
    
    If a > 9 Then
        HexToChar = a - 10 + corLowerAChar
    Else
        HexToChar = a + corZeroChar
    End If
End Function


''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim other As Guid
    
    On Error GoTo Catch
    Set other = Value
'    Equals = Other.InternalEquals(mGuid)
    Equals = IsEqualGUID(mGuid, ByVal other.Handle)

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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant)
    Select Case VariantType(Value)
        Case corByteArray
            FromByteArray Value
        
        Case vbString
            Parse Value
            
        Case Else
            If IsArray(Value) Then _
                Throw Cor.NewArgumentNullException(Environment.GetResourceString(Parameter_Value), Environment.GetResourceString(ArgumentNull_Array))
                
    End Select
End Sub

Friend Sub InitReadOnly(ByRef other As VBGUID)
    mGuid = other
    mIsReadOnly = True
End Sub

Friend Sub InitNew()
    mGuid = CoCreateGuid
End Sub



Friend Function InternalCompare(ByRef g As VBGUID) As Long
    With mGuid
        If Not EqualFields(.Data1, g.Data1, InternalCompare) Then Exit Function
        If Not EqualFields(.Data2, g.Data2, InternalCompare) Then Exit Function
        If Not EqualFields(.Data3, g.Data3, InternalCompare) Then Exit Function
        
        Dim i As Long
        For i = 0 To 7
            If Not EqualFields(.Data4(i), g.Data4(i), InternalCompare) Then Exit Function
        Next i
    End With
End Function

Friend Sub Parse(ByVal s As String)
    If Len(s) = 38 Then
        '{C200E360-38C5-11CE-AE62-08002B2B79EF} or
        '(C200E360-38C5-11CE-AE62-08002B2B79EF)
        Select Case Asc(s)
            Case 123:   If Asc(Right$(s, 1)) <> 125 Then FormatError
            Case 40:    If Asc(Right$(s, 1)) <> 41 Then FormatError
            Case Else:  FormatError
        End Select
        s = Mid$(s, 2, 36)
    End If
    
    Select Case Len(s)
        Case 36
            'C200E360-38C5-11CE-AE62-08002B2B79EF
            If s Like "*[!0-9a-fA-F,-]*" Then FormatError
            If Asc(Mid$(s, 9, 1)) <> 45 Then FormatError
            If Asc(Mid$(s, 14, 1)) <> 45 Then FormatError
            If Asc(Mid$(s, 19, 1)) <> 45 Then FormatError
            If Asc(Mid$(s, 24, 1)) <> 45 Then FormatError
            s = "{" & s & "}"
        Case 32
            'C200E36038C511CEAE6208002B2B79EF
            If s Like "*[!0-9a-fA-F]*" Then FormatError
            s = "{" & Left$(s, 8) & "-" & Mid$(s, 9, 4) & "-" & Mid$(s, 13, 4) & "-" & Mid$(s, 17, 4) & "-" & Mid$(s, 21) & "}"
        
        Case Else
            FormatError
    End Select
    
    On Error GoTo errTrap
    mGuid = GUIDFromString(s)
    Exit Sub
    
errTrap:
    FormatError
End Sub

Private Sub FormatError()
    Throw Cor.NewFormatException("Invalid Guid string format.")
End Sub

Friend Sub FromByteArray(ByRef b As Variant)
    Dim Ptr As Long
    Ptr = CorArray.ArrayPointer(b)
    If Ptr = vbNullPtr Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(Parameter_Value), Environment.GetResourceString(ArgumentNull_Array))
    
    Dim Bytes() As Byte
    SAPtr(Bytes) = Ptr
    CopyMemory mGuid, Bytes(LBound(Bytes)), 16
    SAPtr(Bytes) = vbNullPtr
End Sub

Friend Sub FromParts(ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByRef d() As Byte)
    With mGuid
        .Data1 = a
        .Data2 = b
        .Data3 = c
        
        If CorArray.IsNull(d) Then _
            Throw Cor.NewArgumentNullException("Bytes", Environment.GetResourceString(ArgumentNull_Array))
        If CorArray.Length(d) <> 8 Then _
            Throw Cor.NewArgumentException("Array must be 8 bytes in length.", "Bytes")
        
        CopyMemory .Data4(0), d(LBound(d)), 8
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EqualFields(ByVal MyField As Long, ByVal TheirField As Long, ByRef Result As Long) As Boolean
    If MyField < TheirField Then
        Result = -1
    ElseIf MyField > TheirField Then
        Result = 1
    Else
        Result = 0
        EqualFields = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mGuid.Data1 = .ReadProperty("Data1")
        mGuid.Data2 = .ReadProperty("Data2")
        mGuid.Data3 = .ReadProperty("Data3")
        
        Dim b() As Byte
        b = .ReadProperty("Data4")
        CopyMemory mGuid.Data4(0), b(0), 8
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Data1", mGuid.Data1
        .WriteProperty "Data2", mGuid.Data2
        .WriteProperty "Data3", mGuid.Data3
        .WriteProperty "Data4", mGuid.Data4
    End With
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitArray"
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
' Module: BitArray
'

''
' Manages a compact array of bit values. Each bit represents
' a boolean, where True = 1, and False = 0.
'
' @remarks The default size is 32 bits represented. A <b>BitArray</b> with a
' different size can be creates using the public constructor or static constructors.
' <p>The following are some of the available methods to create a <b>BitArray</b>.
' <pre>
' Set bits = New BitArray
' - or -
' Set bits = NewBitArray(128)
' - or -
' Set bits = BitArray.FromBooleans(BooleanArray)
' </pre></p>
' @see Constructors
' @see BitArrayStatic
'
Option Explicit
Implements IObject
Implements IEnumerable
Implements ICollection
Implements ICloneable

Private Const PROP_UBOUND       As String = "Ubound"
Private Const PROP_LENGTH       As String = "Length"
Private Const PROP_BITSUBOUND   As String = "BitsUBound"
Private Const PROP_BITS         As String = "Bits"

Private mBits()     As Long
Private mLength     As Long
Private mUBound     As Long
Private mVersion    As Long

''
' Performs a bitwise AND on the current instance of BitArray using the
' bits in another instance of BitArray.
'
' @param bits The BitArray object with which to perform the bitwise
' AND operation with.
' @return The internal set of bits is modified based on the operation,
' however, the object returns itself for ease of concatenated operations.
'
Public Function AndBits(ByVal Bits As BitArray) As BitArray
    If Bits Is Nothing Then _
        Throw Cor.NewArgumentNullException("Bits", Environment.GetResourceString(ArgumentNull_Array))
    If Bits.Length <> mLength Then _
        Throw Cor.NewArgumentException("Both arrays must have the same length.", "bits")
    
    Call Bits.ApplyAndTo(mBits)
    mVersion = mVersion + 1
    Set AndBits = Me
End Function

''
' Returns a duplicate of the current instance.
'
' @return A duplicate of the current instance.
'
Public Function Clone() As BitArray
    Set Clone = New BitArray
    Call Clone.CloneHelper(mBits, mLength)
End Function

''
' Copies the internal bit-array to a compatible array.
'
' @param DstArray The array to copy the values to.
' @param Index The starting position in the array to
' begin copying to.
' @remarks CopyTo supports the copying to three array types: Long, Byte
' and Boolean. When copying to Long or Byte type arrays, the bits in the
' elements are equivalents to the bits in the BitArray. When copying to
' a Boolean array, each boolean element represents a single bit in the array.
'
Public Sub CopyTo(ByRef DstArray As Variant, ByVal Index As Long)
    If cArray.IsNull(DstArray) Then _
        Throw Cor.NewArgumentNullException("dstArray", Environment.GetResourceString(ArgumentNull_Array))
    If Index < LBound(DstArray) Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Message:=Environment.GetResourceString(ArgumentOutOfRange_LBound))
    If cArray.GetRank(DstArray) <> 1 Then _
        Throw Cor.NewRankException(Environment.GetResourceString(Rank_MultiDimNotSupported))
    
    Select Case VarType(DstArray)
        Case vbLongArray
            Call cArray.CopyEx(mBits, 0, DstArray, Index, (mLength + 31) \ 32)
        
        Case vbByteArray
            If (Index + (mLength + 7) \ 8) - 1 > UBound(DstArray) Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayPlusOffTooSmall), "index")
            
            Call Buffer.BlockCopy(mBits, 0, DstArray, Index - LBound(DstArray), (mLength + 7) \ 8)
            
        Case vbBooleanArray
            If cArray.GetLength(DstArray) < mLength Then _
                Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_ArrayPlusOffTooSmall), "index")
            
            Dim i As Long
            For i = 0 To mLength - 1
                DstArray(Index + i) = CBool(mBits(i \ 32) And Powers(i Mod 32))
            Next i
        
        Case Else
            Throw Cor.NewArrayTypeMismatchException(Environment.GetResourceString(ArrayTypeMismatch_Incompatible))
    End Select
End Sub

''
' Returns the number of bits being represented in the array.
'
' @return The number of bits represented.
' @remarks This returns the same value as Length, however, this
' property is read-only.
'
Public Property Get Count() As Long
    Count = mLength
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns an enumerator for this instance.
'
' @return An enumerator
'
Public Function GetEnumerator() As IEnumerator
    Dim Ret As New BitArrayEnumerator
    Call Ret.Init(Me)
    Set GetEnumerator = Ret
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns if the instance is a read-only instance.
'
' @return Value indicating if the instance is read-only.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = False
End Property

''
' Returns the value of a specific bit in the array.
'
' @param index The specific bit to return.
' @return Value of the specified bit.
' @remarks The index is zero based.
'
Public Property Get Item(ByVal Index As Long) As Boolean
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))
    
    Dim bit As Long
    Index = GetBitIndex(Index, bit)
    Item = CBool(mBits(Index) And bit)
End Property

''
' Sets the value of a specific bit in the array.
'
' @param index The specific bit to set.
' @param RHS The value to set the specified bit to.
' @remarks The index is zero based.'
'
Public Property Let Item(ByVal Index As Long, ByVal RHS As Boolean)
    If Index < 0 Or Index >= mLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Index", Message:=Environment.GetResourceString(ArgumentOutOfRange_Index))

    Dim bit As Long
    Index = GetBitIndex(Index, bit)
    If RHS Then
        mBits(Index) = mBits(Index) Or bit
    Else
        mBits(Index) = mBits(Index) And (Not bit)
    End If
    mVersion = mVersion + 1
End Property

''
' Returns the number of bits represented in the array.
'
' @return The number of bits represented.
'
Public Property Get Length() As Long
    Length = mLength
End Property

''
' Sets the number of bits represented in the array.
'
' @param RHS The number of bits to represent.
' @remarks If the length of the array is increased, then the
' newly available bits are set to 0 (False).
'
Public Property Let Length(ByVal RHS As Long)
    If RHS <> mLength Then
        Dim Index As Long
        Index = (RHS - 1) \ 32
        ReDim Preserve mBits(0 To Index)
        mUBound = Index
        
        Dim bit As Long
        bit = RHS - Index * 32
        If bit > 0 Then
            bit = Powers(bit - 1)
            mBits(mUBound) = mBits(mUBound) And bit
        End If
        mLength = RHS
    End If
    mVersion = mVersion + 1
End Property

''
' Returns an enumerator to be used in For..Each loops.
'
' @return An enumerator.
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

''
' Performs bitwise negate operation on the internal array.
'
' @return The instance returns a reference to itself to allow
' for concatenated operations.
'
Public Function NotBits() As BitArray
    Dim i As Long
    For i = 0 To mUBound
        mBits(i) = Not mBits(i)
    Next i
    mVersion = mVersion + 1
    Set NotBits = Me
End Function

''
' Performs a bitwise OR on the current instance of BitArray using the
' bits in another instance of BitArray.
'
' @param bits The BitArray object with which to perform the bitwise
' OR operation with.
' @return The internal set of bits is modified based on the operation,
' however, the object returns itself for ease of concatenated operations.
'
Public Function OrBits(ByVal Bits As BitArray) As BitArray
    If Bits Is Nothing Then _
        Throw Cor.NewArgumentNullException("Bits", Environment.GetResourceString(ArgumentNull_Array))
    If Bits.Length <> mLength Then _
        Throw Cor.NewArgumentException("Both arrays must have the same length.", "bits")
    
    Call Bits.ApplyOrTo(mBits)
    mVersion = mVersion + 1
    Set OrBits = Me
End Function

''
' Sets all of the bits to the specified value of True (1) or False (0).
'
' @param value The value to set all the bits to.
'
Public Sub SetAll(ByVal Value As Boolean)
    Dim i As Long
    For i = 0 To mUBound
        mBits(i) = Value
    Next i
    mVersion = mVersion + 1
End Sub

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Performs a bitwise XOR on the current instance of BitArray using the
' bits in another instance of BitArray.
'
' @param bits The BitArray object with which to perform the bitwise
' XOR operation with.
' @return The internal set of bits is modified based on the operation,
' however, the object returns itself for ease of concatenated operations.
'
Public Function XorBits(ByVal Bits As BitArray) As BitArray
    If Bits Is Nothing Then _
        Throw Cor.NewArgumentNullException("Bits", Environment.GetResourceString(ArgumentNull_Array))
    If Bits.Length <> mLength Then _
        Throw Cor.NewArgumentException("Both arrays must have the same length.", "bits")

    Call Bits.ApplyXorTo(mBits)
    mVersion = mVersion + 1
    Set XorBits = Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub ApplyAndTo(ByRef Bits() As Long)
    Dim i As Long
    For i = 0 To mUBound
        Bits(i) = Bits(i) And mBits(i)
    Next i
End Sub

Friend Sub ApplyOrTo(ByRef Bits() As Long)
    Dim i As Long
    For i = 0 To mUBound
        Bits(i) = Bits(i) Or mBits(i)
    Next i
End Sub

Friend Sub ApplyXorTo(ByRef Bits() As Long)
    Dim i As Long
    For i = 0 To mUBound
        Bits(i) = Bits(i) Xor mBits(i)
    Next
End Sub

Friend Sub CloneHelper(ByRef Bits() As Long, ByVal Length As Long)
    mBits = Bits
    mLength = Length
End Sub

Friend Sub Init(ByVal Length As Long, ByVal Value As Boolean)
    Me.Length = Length
    Call SetAll(Value)
End Sub

Friend Sub InitFromBooleans(ByRef Values() As Boolean)
    If cArray.IsNull(Values) Then _
        Throw Cor.NewArgumentNullException("Values", Environment.GetResourceString(ArgumentNull_Array))

    Dim i   As Long
    Dim lb  As Long
    Me.Length = cArray.GetLength(Values, 1)
    lb = LBound(Values)
    For i = 0 To mLength - 1
        Item(i) = Values(lb + i)
    Next i
End Sub

Friend Sub InitFromBytes(ByRef Values() As Byte)
    If cArray.IsNull(Values) Then _
        Throw Cor.NewArgumentNullException("Values", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim Size As Long
    Size = cArray.GetLength(Values, 1)
    Me.Length = Size * 8
    Call CopyMemory(mBits(0), Values(0), Size)
End Sub

Friend Sub InitFromLongs(ByRef Values() As Long)
    If cArray.IsNull(Values) Then _
        Throw Cor.NewArgumentNullException("Values", Environment.GetResourceString(ArgumentNull_Array))
    
    mBits = Values
    mUBound = UBound(Values)
    mLength = (mUBound + 1) * 32
End Sub

Friend Property Get Version() As Long
    Version = mVersion
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    ReDim mBits(0)
    mLength = 32
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mUBound = .ReadProperty(PROP_UBOUND)
        mLength = .ReadProperty(PROP_LENGTH)
        
        ReDim mBits(0 To CLng(.ReadProperty(PROP_BITSUBOUND)))
        Dim i As Long
        For i = 0 To UBound(mBits)
            mBits(i) = .ReadProperty(PROP_BITS & i, 0)
        Next i
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_UBOUND, mUBound)
        Call .WriteProperty(PROP_LENGTH, mLength)
        Call .WriteProperty(PROP_BITSUBOUND, UBound(mBits))
        
        Dim i As Long
        For i = 0 To UBound(mBits)
            Call .WriteProperty(PROP_BITS & i, mBits(i))
        Next i
    End With
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

Private Function GetBitIndex(ByVal Index As Long, ByRef bit As Long) As Long
    GetBitIndex = Index \ 32
    bit = Powers(Index - GetBitIndex * 32)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICollection Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long)
    Call CopyTo(Arr, Index)
End Sub

Private Property Get ICollection_Count() As Long
    ICollection_Count = Count
End Property

Private Function ICollection_GetEnumerator() As IEnumerator
    Set ICollection_GetEnumerator = GetEnumerator
End Function

Private Function ICollection_NewEnum() As stdole.IUnknown
    Set ICollection_NewEnum = NewEnum
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IEnumerable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Function


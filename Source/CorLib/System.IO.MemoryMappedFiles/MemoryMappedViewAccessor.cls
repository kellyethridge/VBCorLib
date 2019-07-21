VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryMappedViewAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: MemoryMappedViewAccessor
'

''
' Represents a randomly accessed view of a memory-mapped file.
'
' @remarks Use the <a href="MemoryMappedFile.CreateViewAccessor.html">CreateViewAccessor</a> method of a <a href="MemoryMappedFile.html">MemoryMappedFile</a> object to get this view.
Option Explicit
Implements IObject

Private mMappingHandle  As SafeFileHandle
Private mViewAddress    As Long
Private mStartIndex     As Long
Private mViewData()     As Byte
Private mViewDataSA     As SafeArray1d
Private mCanWrite       As Boolean
Private mCapacity       As Long
Private mOffset         As Currency


''
' Determines whether the accessory is writable.
'
' @return <b>True</b> if the accessor is writable; otherwise, <b>False</b>.
Public Property Get CanWrite() As Boolean
    CanWrite = mCanWrite
End Property

''
' Determines whether the accessor is readable.
'
' @return <b>True</b> if the accessor is readable; otherwise, <b>False</b>.
' @remarks This will always return <b>True</b> if the accessor is successfully created.
Public Property Get CanRead() As Boolean
    CanRead = True
End Property

''
' Gets the capacity of the accessor.
'
' @return The capacity of the accessor.
Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

''
' Gets the number of bytes by which the starting position of this view is offset from the beginning of the memory-mapped file.
'
' @return The number of bytes between the starting position of this view and the beginning of the memory-mapped file.
Public Property Get PointerOffset() As Currency
    PointerOffset = mOffset
End Property

''
' Reads a byte value from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' <h4>Exceptions</h4>
' @remarks <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadByte(ByVal Position As Long) As Byte
    Position = GetSafeIndex(Position, vbSizeOfByte)
    ReadByte = mViewData(Position)
End Function

''
' Reads a boolean value from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadBoolean(ByVal Position As Long) As Boolean
    ReadBoolean = ReadByte(Position) <> 0
End Function

''
' Reads a character from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks <p>The <b>ReadChar</b> method does not attempt to apply any encoding. It will simply read two consecutive bytes as a character.</p>
' <p>This method will return the same value as <a href="MemoryMappedViewAccessor.ReadInt16.html">MemoryMappedViewAccessor.ReadInt16</a>.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadChar(ByVal Position As Long) As Integer
    ReadChar = ReadInt16(Position)
End Function

''
' Reads a 16-bit integer from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>

Public Function ReadInt16(ByVal Position As Long) As Integer
    Position = GetSafeIndex(Position, vbSizeOfInteger)
    ReadInt16 = AsWord(mViewData(Position))
End Function

''
' Reads a 32-bit integer from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadInt32(ByVal Position As Long) As Long
    Position = GetSafeIndex(Position, vbSizeOfLong)
    ReadInt32 = AsLong(mViewData(Position))
End Function

''
' Reads a 64-bit integer from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadInt64(ByVal Position As Long) As Int64
    Position = GetSafeIndex(Position, vbSizeOfInt64)
    AssignQWord mViewData(Position), ReadInt64
End Function

''
' Reads a 64-bit currency from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks <p>The <b>ReadCurrency</b> method reads 64-bits but does not scale the value
' to create a <b>Currency</b> without a fractional part. For example if the 64-bits read
' is equal to 1 then the returned currency value would be 0.0001.</p>
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadCurrency(ByVal Position As Long) As Currency
    Position = GetSafeIndex(Position, vbSizeOfCurrency)
    ReadCurrency = AsCurr(mViewData(Position))
End Function

''
' Reads a double-precision floating-point value from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadDouble(ByVal Position As Long) As Double
    Position = GetSafeIndex(Position, vbSizeOfDouble)
    ReadDouble = AsDouble(mViewData(Position))
End Function

''
' Reads a single-precision floating-point value from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadSingle(ByVal Position As Long) As Single
    Position = GetSafeIndex(Position, vbSizeOfSingle)
    ReadSingle = AsSingle(mViewData(Position))
End Function

''
' Reads a decimal value from the accessor.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @return The value that was read.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to read a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadDecimal(ByVal Position As Long) As Variant
    Position = GetSafeIndex(Position, vbSizeOfDecimal)
    ReadDecimal = BitConverter.ToDecimal(mViewData, Position)
End Function

''
' Reads a structure of the type passed in from the accessor into a provided reference.
'
' @param Position The number of bytes into the accessor at which to begin reading.
' @param Structure The structure to contain the read data.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td>
'       <td>
'           There are not enough bytes after <i>Position</i> to read a value.<br/>
'           -or-<br/>
'           <i>Structure</i> is not a user-defined or primitive type.
'       </td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Sub Read(ByVal Position As Long, ByRef Structure As Variant)
    Dim Record      As IRecordInfo
    Dim RecordPtr   As Long
    Dim Size        As Long
    Dim Ptr         As Long
    
    Select Case VarType(Structure)
        Case vbBoolean
            Structure = ReadBoolean(Position)
        Case vbByte
            Structure = ReadByte(Position)
        Case vbInteger
            Structure = ReadInt16(Position)
        Case vbLong
            Structure = ReadInt32(Position)
        Case vbCurrency
            Structure = ReadCurrency(Position)
        Case vbSingle
            Structure = ReadSingle(Position)
        Case vbDouble
            Structure = ReadDouble(Position)
        Case vbDate
            Structure = CorDateTime.GetOADate(ReadCurrency(Position))
        Case vbDecimal
            Structure = ReadDecimal(Position)
        Case vbUserDefinedType
            RecordPtr = MemLong(VarPtr(Structure) + 12)
            ObjectPtr(Record) = RecordPtr
            
            On Error GoTo Catch
            ValidateStructure Record.GetTypeInfo
            Size = Record.GetSize
            Position = GetSafeIndex(Position, Size)
            Ptr = MemLong(VarPtr(Structure) + VARIANTDATA_OFFSET)
            CopyMemory ByVal Ptr, mViewData(Position), Size
        Case Else
            Error.Argument Argument_InvalidStructure, "Structure"
    End Select
    
    GoSub Finally
    Exit Sub
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    ObjectPtr(Record) = vbNullPtr
    Return
End Sub

''
' Reads structures of the type passed in from the accessor into an array of the type passed in.
'
' @param Position The number of bytes in the accessor at which to begin reading.
' @param Arr The array to contain the structures read from the accessor.
' @param Offset The index in array in which to place the first copied structure.
' @param Count The number of structures of the type passed in to read from the accessor.
' @return The number of structures read into <i>Arr</i>. This value can be less than <i>Count</i> if there are fewer structures available, or zero if the end of the accessor is reached.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Arr</i> is not large enough to contain <i>Count</i> of structures (starting from <i>Position</i>).</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Arr</i> is <b>null</b></td>
'    </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
' </table>
Public Function ReadArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long) As Long
    Dim Ptr         As Long
    Dim pvData      As Long
    Dim SizeOfType  As Long
    
    ValidateArrayRange Arr, Offset, Count, NameOfOffset
    
    Select Case VarType(Arr) And &HFF
        Case vbBoolean
            ReadArray = ReadBooleanArray(Position, Arr, Offset, Count)
            Exit Function
        Case vbByte
            SizeOfType = vbSizeOfByte
        Case vbInteger
            SizeOfType = vbSizeOfInteger
        Case vbLong
            SizeOfType = vbSizeOfLong
        Case vbCurrency
            SizeOfType = vbSizeOfCurrency
        Case vbDate
            ReadArray = ReadDateArray(Position, Arr, Offset, Count)
            Exit Function
        Case vbSingle
            SizeOfType = vbSizeOfSingle
        Case vbDouble
            SizeOfType = vbSizeOfDouble
        Case vbUserDefinedType
            Dim Record As IRecordInfo
            
            Ptr = SAPtrV(Arr)
            Set Record = SafeArrayGetRecordInfo(Ptr)
            ValidateStructure Record.GetTypeInfo
            SizeOfType = Record.GetSize
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_TypeNotSupported, TypeName(Arr)), "Arr")
    End Select
    
    Count = CorMath.Min(Count, (mCapacity - Position) \ SizeOfType)
    Position = GetSafeIndex(Position, Count * SizeOfType)
    Ptr = SAPtrV(Arr)
    pvData = MemLong(Ptr + PVDATA_OFFSET) + Offset * SizeOfType
    CopyMemory ByVal pvData, mViewData(Position), Count * SizeOfType
    ReadArray = Count
End Function

' VB Boolean is two bytes, however to match how .NET deals with
' reading booleans through the accessor, we only use one
' byte to represent the boolean value. We also want the final
' boolean value to be a true value, meaning the integer value will
' be either 0 or -1. If we just used CopyMemory we could get values
' other than those shoved into the boolean variable.
Private Function ReadBooleanArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long) As Long
    Dim i As Long
    
    Count = CorMath.Min(Count, mCapacity - Position)
    Position = GetSafeIndex(Position, Count)
    
    For i = 0 To Count - 1
        Arr(Offset + i) = CBool(mViewData(Position + i))
    Next
    
    ReadBooleanArray = Count
End Function

' VB Date is similar to a Double. A Date stores the days portion as the
' whole number and the time portion as the fractional portion. However,
' .NET stores dates as a 64-bit integer representing the number of ticks
' from 1/1/0001. We maintain the same representation when reading and
' writing through the accessor.
Private Function ReadDateArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long) As Long
    Dim Milliseconds As Currency
    Dim i As Long
    
    Count = CorMath.Min(Count, (mCapacity - Position) \ vbSizeOfCurrency)
    Position = GetSafeIndex(Position, Count * vbSizeOfCurrency)
    
    For i = 0 To Count - 1
        Milliseconds = AsCurr(mViewData(Position + i * vbSizeOfCurrency))
        Arr(Offset + i) = CorDateTime.GetOADate(Milliseconds)
    Next
    
    ReadDateArray = Count
End Function

''
' Writes the value to the accessor.
'
' @param Position The number of bytes into the accessor at which to begin writing.
' @param Value The value to write.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes after <i>Position</i> to write a value.</td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Position</i> is less than zero or greater than the capacity of the accessor.</td>
'   </tr>
'   <tr>
'       <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The accessor does not support writing.</td>
'   </tr>
' </table>
Public Sub WriteValue(ByVal Position As Long, ByRef Value As Variant)
    If Not CanWrite Then _
        Error.NotSupported NotSupported_Writing
    
    Select Case VarType(Value)
        Case vbBoolean
            ' .NET writes out booleans as one byte using 0 or 1 to represent the value.
            Position = GetSafeIndex(Position, vbSizeOfByte)
            mViewData(Position) = IIfLong(Value, 1, 0)
        Case vbByte
            Position = GetSafeIndex(Position, vbSizeOfByte)
            mViewData(Position) = Value
        Case vbInteger
            Position = GetSafeIndex(Position, vbSizeOfInteger)
            AsWord(mViewData(Position)) = Value
        Case vbLong
            Position = GetSafeIndex(Position, vbSizeOfLong)
            AsLong(mViewData(Position)) = Value
        Case vbCurrency
            Position = GetSafeIndex(Position, vbSizeOfCurrency)
            AsCurr(mViewData(Position)) = Value
        Case vbDate
            ' .NET stores dates as 64-bit integers. we do the same thing here.
            Position = GetSafeIndex(Position, vbSizeOfCurrency)
            AsCurr(mViewData(Position)) = CorDateTime.GetTotalMilliseconds(Value)
        Case vbSingle
            Position = GetSafeIndex(Position, vbSizeOfSingle)
            AsSingle(mViewData(Position)) = Value
        Case vbDouble
            Position = GetSafeIndex(Position, vbSizeOfDouble)
            AsDouble(mViewData(Position)) = Value
        Case vbDecimal
            WriteDecimal Position, Value
        Case vbUserDefinedType
            WriteUdt Position, Value
        Case Else
            Error.Argument Argument_NeedIntrinsicType
    End Select
End Sub

Private Sub WriteDecimal(ByVal Position As Long, ByRef Value As Variant)
    Dim Data() As Byte
    
    Position = GetSafeIndex(Position, vbSizeOfDecimal)
    Data = BitConverter.GetBytes(Value)
    CopyMemory mViewData(Position), Data(0), vbSizeOfDecimal
End Sub

Private Sub WriteUdt(ByVal Position As Long, ByRef Value As Variant)
    Dim Record  As IRecordInfo
    Dim Size    As Long
    Dim Ptr     As Long
    
    ObjectPtr(Record) = MemLong(VarPtr(Value) + VARIANTRECORD_OFFSET)
    
    On Error GoTo Catch
    ValidateStructure Record.GetTypeInfo
    Size = Record.GetSize
    Position = GetSafeIndex(Position, Size)
    Ptr = MemLong(VarPtr(Value) + VARIANTDATA_OFFSET)
    CopyMemory mViewData(Position), ByVal Ptr, Size
    Exit Sub
    
Catch:
    GoSub Finally
    ThrowOrErr Err
Finally:
    ObjectPtr(Record) = vbNullPtr
    Return
End Sub

''
' Writes structures of the type from an array passed into the accessor.
'
' @param The number of bytes into the accessor at which to begin writing.
' @param The array to write into the accessor.
' @param The index in <i>Arr</i> to start writing from.
' @param The number of structures in <i>Arr</i> to write.
' @remarks
' <h4>Exceptions</h4>
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td><td>There are not enough bytes in the accessor after <i>Position</i> to write the number of structures specified by <i>Count</i>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>Position</i> is less than zero or greater than the capacity of the accessor.<br/>
'           -or-<br/>
'           <i>Offset</i> is less than the lower-bound of <i>Arr</i>.<br/>
'           -or-<br/>
'           <i>Count</i> is less than zero.
'        </td>
'    </tr>
'    <tr>
'        <td><a href="NotSupportedException.html">NotSupportedException</a></td><td>The accessor does not support writing.</td>
'    </tr>
' </table>
Public Sub WriteArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long)
    Dim Ptr         As Long
    Dim pvData      As Long
    Dim SizeOfType  As Long
    Dim DataType    As VbVarType
    
    ValidateArrayRange Arr, Offset, Count, NameOfOffset
    
    If Not CanWrite Then _
        Error.NotSupported NotSupported_Writing
    If Position < 0 Then _
        Error.ArgumentOutOfRange "Position", ArgumentOutOfRange_NeedNonNegNum
    
    DataType = VarType(Arr) And &HFF
    
    Select Case DataType
        Case vbBoolean
            WriteBooleanArray Position, Arr, Offset, Count
            Exit Sub
        Case vbDate
            WriteDateArray Position, Arr, Offset, Count
            Exit Sub
        Case vbVariant
            WriteVariantArray Position, Arr, Offset, Count
            Exit Sub
        Case vbByte, vbInteger, vbLong, vbCurrency, vbSingle, vbDouble
            SizeOfType = SizeOf(DataType)
        Case vbUserDefinedType
            Dim Record As IRecordInfo
            
            Ptr = SAPtrV(Arr)
            Set Record = SafeArrayGetRecordInfo(Ptr)
            ValidateStructure Record.GetTypeInfo
            SizeOfType = Record.GetSize
        Case Else
            Throw Cor.NewArgumentException(Environment.GetResourceString(Arg_TypeNotSupported, TypeName(Arr)), "Arr")
    End Select
    
    Position = GetSafeIndex(Position, Count * SizeOfType)
    Ptr = SAPtrV(Arr)
    pvData = MemLong(Ptr + PVDATA_OFFSET) + Offset * SizeOfType
    CopyMemory mViewData(Position), ByVal pvData, Count * SizeOfType
End Sub

Private Sub WriteBooleanArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long)
    Dim i As Long
    
    Position = GetSafeIndex(Position, Count)
    
    For i = 0 To Count - 1
        ' .NET writes out booleans as one byte using 0 or 1 to represent the value.
        mViewData(Position + i) = IIfLong(Arr(Offset + i), 1, 0)
    Next
End Sub

Private Sub WriteDateArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long)
    Dim pvData  As Long
    Dim i       As Long
    
    Position = GetSafeIndex(Position, Count * vbSizeOfDate)
    pvData = MemLong(SAPtrV(Arr) + PVDATA_OFFSET) + Offset * vbSizeOfDate
    
    For i = 0 To Count - 1
        AsCurr(mViewData(Position + i * vbSizeOfDate)) = CorDateTime.GetTotalMilliseconds(Arr(Offset + i))
    Next
End Sub

' A variant array can hold any datatype in each element, so we will iterate through
' each element and try to write them out individually. Each element will be validated
' for a supported type to write. If an error occurs, the state of the view is unknown.
Private Sub WriteVariantArray(ByVal Position As Long, ByRef Arr As Variant, ByVal Offset As Long, ByVal Count As Long)
    Dim i As Long
    
    For i = 0 To Count - 1
        WriteValue Position, Arr(Offset + i)
        Position = Position + GetSizeOfType(Arr(Offset + i))
    Next
End Sub

''
' Clears all buffers for this view and causes any buffered data to be written to the underlying file.
Public Sub Flush()
    If FlushViewOfFile(mViewAddress, 0) = BOOL_FALSE Then
        Error.Win32Error Err.LastDllError
    End If
End Sub

''
' Determines whether the specified value passed in is equal to the current object.
'
' @param Value The value to compare to the current object.
' @return <b>True</b> if the specified value is equal to the current object; otherwise, <b>False</b>.
' @see IObject
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Serves as the default hash function.
'
' @return A hash code for the current object.
' @see IObject
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string that represents the current object.
'
' @return A string the represents the current object.
' @see IObject
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal MappingHandle As SafeFileHandle, ByVal Offset As Currency, ByVal Size As Long, ByVal Access As MemoryMappedFileAccess)
    Dim PageAlignedOffset As Currency
    Dim Info As SYSTEM_INFO
    
    Set mMappingHandle = MappingHandle
    mOffset = Int(Offset)
    GetSystemInfo Info
    
    ' MapViewOfFile requires the offset into a file to be virtual page aligned, so we will
    ' adjust for it and simply hide the begining of the allocated array that was not intended
    ' to be part of the view by using a starting index into the view data.
    mStartIndex = CLng(Int(mOffset Mod CCur(Info.dwAllocationGranularity)))
    
    If mStartIndex = 0 Then
        ' the offset is already page aligned so no adjustment needed.
        PageAlignedOffset = mOffset
    Else
        ' calculate the number of whole pages into the file to be used
        ' as the new offset for the view. The mStartIndex will be used to skip
        ' any portion that is in the view data that should be hidden.
        PageAlignedOffset = Int(mOffset / Info.dwAllocationGranularity)
    End If
    
    With CInt64(PageAlignedOffset)
        mViewAddress = MapViewOfFile(mMappingHandle.DangerousGetHandle, GetDesiredAccess(Access), .HighInt32, .LowInt32, Size)
    End With
    
    If mViewAddress = vbNullPtr Then _
        Error.Win32Error Err.LastDllError
        
    SetCapacity Size
    InitViewData
End Sub

Private Sub Class_Terminate()
    If mViewAddress <> vbNullPtr Then
        ReleaseViewData
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCapacity(ByVal Size As Long)
    Dim Info As MEMORY_BASIC_INFORMATION
    
    If Size = 0 Then
        ' if no size is set then we use the size of the mapped region
        ' rounded up to the next virtual page boundary. For instance
        ' if the system virtual page size is 4096 bytes, then the region size
        ' will be the next even 4096 byte page boundary after the size of
        ' mapped file, so a file that is 6000 bytes will create a region
        ' size of 8192 bytes.
        VirtualQuery ByVal mViewAddress, Info, LenB(Info)
        Size = Info.RegionSize
    End If

    mCapacity = Size
End Sub

Private Sub InitViewData()
    With mViewDataSA
        .cbElements = 1
        .cDims = 1
        .cElements = mCapacity
        .pvData = mViewAddress
    End With
    
    SAPtr(mViewData) = VarPtr(mViewDataSA)
End Sub

Private Sub ReleaseViewData()
    SAPtr(mViewData) = vbNullPtr
    UnmapViewOfFile mViewAddress
End Sub

Private Function GetDesiredAccess(ByVal Access As MemoryMappedFileAccess) As Long
    Select Case Access
        Case MemoryMappedFileAccess.ReadWrite
            GetDesiredAccess = FILE_MAP_ALL_ACCESS
            mCanWrite = True
        Case MemoryMappedFileAccess.ReadOnly
            GetDesiredAccess = FILE_MAP_READ
        Case MemoryMappedFileAccess.CopyOnWrite
            GetDesiredAccess = FILE_MAP_COPY
            mCanWrite = True
        Case MemoryMappedFileAccess.ReadExecute
            GetDesiredAccess = FILE_MAP_READ Or FILE_MAP_EXECUTE
        Case MemoryMappedFileAccess.ReadWriteExecute
            GetDesiredAccess = FILE_MAP_ALL_ACCESS Or FILE_MAP_EXECUTE
            mCanWrite = True
        Case Else
            Error.ArgumentOutOfRange "Access"
    End Select
End Function

Private Function GetSafeIndex(ByVal Position As Long, ByVal SizeOfType As Long) As Long
    Dim AdjustedPosition As Long
    
    If Position < 0 Then _
        Error.ArgumentOutOfRange "Position", ArgumentOutOfRange_NeedNonNegNum
    
    AdjustedPosition = mStartIndex + Position
    
    If AdjustedPosition > mCapacity - SizeOfType Then
        If AdjustedPosition >= mCapacity Then _
            Error.ArgumentOutOfRange "Position", ArgumentOutOfRange_PositionLessThanCapacityRequired
        
        Error.Argument Argument_NotEnoughBytesToRead, "Position"
    End If
    
    GetSafeIndex = AdjustedPosition
End Function

Private Sub ValidateStructure(ByVal Info As ITypeInfo)
    Dim Attr    As TYPEATTR
    Dim AttrPtr As Long
    Dim desc    As VARDESC
    Dim DescPtr As Long
    Dim i       As Long
    
    Info.GetTypeAttr AttrPtr
    CopyMemory Attr, ByVal AttrPtr, LenB(Attr)
    Info.ReleaseTypeAttr AttrPtr
    
    For i = 0 To Attr.cVars - 1
        Info.GetVarDesc i, DescPtr
        CopyMemory desc, ByVal DescPtr, LenB(desc)
        Info.ReleaseVarDesc DescPtr
        
        If desc.elemdescVar.tdesc.vt = VT_USERDEFINEDTYPE Then
            Dim InnerInfo As ITypeInfo
            Info.GetRefTypeInfo desc.elemdescVar.tdesc.hRefType, InnerInfo
            ValidateStructure InnerInfo
        Else
            If Not IsValidDataType(desc.elemdescVar.tdesc.vt) Then
                Error.Argument Argument_InvalidStructure, "Structure"
            End If
        End If
    Next
End Sub

Private Function IsValidDataType(ByVal DataType As VbVarType) As Boolean
    Select Case DataType
        Case vbBoolean, vbByte, vbInteger, vbLong, vbCurrency, vbSingle, vbDouble, vbDate, vbUserDefinedType, VT_CARRAY
            IsValidDataType = True
    End Select
End Function

Private Function GetSizeOfType(ByRef Value As Variant) As Long
    Select Case VarType(Value)
        Case vbByte
            GetSizeOfType = 1
        Case Else
            GetSizeOfType = SizeOf(VarType(Value))
    End Select
End Function


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

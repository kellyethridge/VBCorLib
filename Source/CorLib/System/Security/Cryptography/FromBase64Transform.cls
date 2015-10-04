VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromBase64Transform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2015 Kelly Ethridge
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
' Module: FromBase64Transform
'

''
' Transforms data from base-64 to plain text.
'
' @remarks This class is used primarily in conjunction with the <b>CryptoStream</b> class and shouldn't
' be used to process large amounts of data manually. For that use the <b>Convert.FromBase64String</b> method.
'
Option Explicit
Implements IObject
Implements ICryptoTransform

Private Const Base64TerminatorChar As Long = corEqualChar


''
' Represents choices to ignore white spaces when
' tranforming blocks of text.
'
' @param IgnoreWhiteSpaces Any whitespace is skipped and the processing
' begins with the next non-whitespace character.
' @param DoNotIgnoreWhiteSpaces If whitespace is encountered during
' processing, then an exception will be thrown.
' @remarks Whitespaces are defined as characters with the ascii value of 32, 9, 10, 11, 12, 13, 133, 160.
'
Public Enum FromBase64TransformMode
    IgnoreWhiteSpaces = 0
    DoNotIgnoreWhiteSpaces = 1
End Enum


Private mIgnoreWhiteSpaces  As Boolean
Private mBits               As Long
Private mBitCount           As Long
Private mTermCount          As Long



''
' Returns if the transform instance can be reused after transforming the final data.
'
' @return Always returns True.
'
Public Property Get CanReuseTransform() As Boolean
    CanReuseTransform = True
End Property

''
' Returns if TransformBlock can transform multiple blocks can be transformed in a single call.
'
' @return Always returns False.
'
Public Property Get CanTransformMultipleBlocks() As Boolean
    CanTransformMultipleBlocks = False
End Property

''
' Returns then block size the input data must be, or be a multiple of.
'
' @return The input block size.
'
Public Property Get InputBlockSize() As Long
    InputBlockSize = 1
End Property

''
' Returns the block size of the output data.
'
' @return The output block size.
'
Public Property Get OutputBlockSize() As Long
    OutputBlockSize = 3
End Property

''
' Releases any resources held in the transform.
'
Public Sub Clear()
    ' here for consistency
End Sub

''
' Transforms a block of data from a base-64 encoding to plain text.
'
' @param InputBuffer The data to be transformed.
' @param InputOffset The starting position in the array to begin transforming.
' @param InputCount The number of bytes to be transformed.
' @param OutputBuffer The array to place the transformed data in.
' @param OutputOffset The starting position to begin placing the output data.
' @return The number of bytes transformed.
'
Public Function TransformBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, ByRef OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    If SAPtr(InputBuffer) = vbNullPtr Then _
        Error.ArgumentNull "InputBuffer", ArgumentNull_Array
    If InputOffset < LBound(InputBuffer) Then _
        Error.ArgumentOutOfRange "InputOffset", ArgumentOutOfRange_ArrayLB
    If InputCount < 0 Then _
        Error.ArgumentOutOfRange "InputCount", ArgumentOutOfRange_NeedNonNegNum
    If InputOffset + InputCount - 1 > UBound(InputBuffer) Then _
        Error.Argument Argument_InvalidOffLen
    If SAPtr(OutputBuffer) = vbNullPtr Then _
        Error.ArgumentNull "OutputBuffer", ArgumentNull_Array
    If OutputOffset < LBound(OutputBuffer) Then _
        Error.ArgumentOutOfRange "OutputOffset", ArgumentOutOfRange_ArrayLB
    
    Dim StartingOffset As Long
    StartingOffset = OutputOffset
    
    Dim OutputBufferUB As Long
    OutputBufferUB = UBound(OutputBuffer)
    
    Do While InputCount > 0
        Dim b As Long
        b = InputBuffer(InputOffset)
        If CanProcessChar(b) Then
            If b <> Base64TerminatorChar Then
                If mTermCount > 0 Then _
                    Error.Format Format_InvalidBase64Character
                    
                b = Base64CharToBits(b)
                If b = corInvalidChar Then _
                    Error.Format Format_InvalidBase64Character
                    
                mBits = mBits Or b
            Else
                mTermCount = mTermCount + 1
                If mTermCount > 2 Then _
                    Error.Format Format_InvalidBase64Character
                
            End If
            mBitCount = mBitCount + 6
            
            If mBitCount <> 24 Then
                mBits = mBits * &H40&
            Else
                Dim NewOffset As Long
                NewOffset = OutputOffset + (2 - mTermCount)
                If NewOffset > OutputBufferUB Then
                    SmallBufferError "OutputBuffer"
                End If
                
                OutputBuffer(OutputOffset) = (mBits And &HFF0000) \ &H10000
                
                Select Case mTermCount
                    Case 0
                        OutputBuffer(OutputOffset + 2) = mBits And &HFF&
                        OutputBuffer(OutputOffset + 1) = (mBits And &HFF00&) \ &H100&

                    Case 1
                        OutputBuffer(OutputOffset + 1) = (mBits And &HFF00&) \ &H100&

                End Select
                
                OutputOffset = NewOffset + 1
                Reset
            End If
        End If
        
        InputOffset = InputOffset + 1
        InputCount = InputCount - 1
    Loop
    
    TransformBlock = OutputOffset - StartingOffset
End Function

''
' Transforms a block of data and any data that has been buffered from previous TransformBlock calls.
'
' @param InputBuffer The remaining data to be transformed.
' @param InputOffset The starting index to being transforming from.
' @param InputCount The number of bytes to transform.
' @return The final transformed data.
'
Public Function TransformFinalBlock(ByRef InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    If SAPtr(InputBuffer) = vbNullPtr Then _
        Error.ArgumentNull "InputBuffer", ArgumentNull_Array
    If InputOffset < LBound(InputBuffer) Then _
        Error.ArgumentOutOfRange "InputOffset", ArgumentOutOfRange_ArrayLB
    If InputCount < 0 Then _
        Error.ArgumentOutOfRange "InputCount", ArgumentOutOfRange_NeedNonNegNum
    If InputOffset + InputCount - 1 > UBound(InputBuffer) Then _
        Error.Argument Argument_InvalidOffLen

    Dim Result() As Byte
    If (mBitCount \ 6) + InputCount < 4 Then
        Result = Cor.NewBytes()
    ElseIf mTermCount > 0 Then
        ReDim Ret(0 To 1 - mTermCount)
        TransformBlock InputBuffer, InputOffset, InputCount, Result, 0
    Else
        ReDim Result(0 To InputCount)
        Dim Count As Long
        
        Count = TransformBlock(InputBuffer, InputOffset, InputCount, Result, 0)
        If Count > 0 Then
            ReDim Preserve Result(0 To Count - 1)
        Else
            Result = Cor.NewBytes()
        End If
    End If
        
    Reset
    TransformFinalBlock = Result
End Function

''
' This function determines if the value passed in is the same
' as the current object instance. Meaning, are the Value and
' this object the same object in memory.
'
' @param Value The value to test for equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a psuedo-unique number used to help identify this
' object in memory. The current method is to return the value
' obtained from ObjPtr. If a different method needs to be impelmented
' then change the method here in this function.
'
' An override might be necessary if the hashcode should be
' derived from a value contained within the class.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
     mIgnoreWhiteSpaces = True
End Sub

Friend Sub Init(ByVal WhiteSpaces As FromBase64TransformMode)
    mIgnoreWhiteSpaces = (WhiteSpaces = IgnoreWhiteSpaces)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByVal ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), ParamName)
End Sub

Private Function CanProcessChar(ByVal Ch As Integer) As Boolean
    If Char.IsWhiteSpaceChar(Ch) Then
        If Not mIgnoreWhiteSpaces Then
            Error.Format Format_InvalidBase64Character
        End If
    Else
        CanProcessChar = True
    End If
End Function

Private Sub Reset()
    mTermCount = 0
    mBits = 0
    mBitCount = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ICryptoTransform_CanReuseTransform() As Boolean
    ICryptoTransform_CanReuseTransform = CanReuseTransform
End Property

Private Property Get ICryptoTransform_CanTransformMultipleBlocks() As Boolean
    ICryptoTransform_CanTransformMultipleBlocks = CanTransformMultipleBlocks
End Property

Private Property Get ICryptoTransform_InputBlockSize() As Long
    ICryptoTransform_InputBlockSize = InputBlockSize
End Property

Private Property Get ICryptoTransform_OutputBlockSize() As Long
    ICryptoTransform_OutputBlockSize = OutputBlockSize
End Property

Private Function ICryptoTransform_TransformBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long, OutputBuffer() As Byte, ByVal OutputOffset As Long) As Long
    ICryptoTransform_TransformBlock = TransformBlock(InputBuffer, InputOffset, InputCount, OutputBuffer, OutputOffset)
End Function

Private Function ICryptoTransform_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    ICryptoTransform_TransformFinalBlock = TransformFinalBlock(InputBuffer, InputOffset, InputCount)
End Function


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

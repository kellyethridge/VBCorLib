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
'    CopyRight (c) 2006 Kelly Ethridge
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
'    Module: FromBase64Transform
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

Private Const DEF_IGNOREWHITESPACES             As Boolean = True
Private Const vbBase64TerminatorChar            As Long = corEqualChar


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
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InputBuffer), InputOffset, InputCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InputBuffer", InputOffset, "InputOffset", InputCount, "InputCount", False)
    
    If SAPtr(OutputBuffer) = vbNullPtr Then _
        Throw Cor.NewArgumentNullException("OutputBuffer", Environment.GetResourceString(ArgumentNull_Array))
    If OutputOffset < LBound(OutputBuffer) Then _
        Throw Cor.NewArgumentOutOfRangeException("OutputOffset", Message:=Environment.GetResourceString(ArgumentOutOfRange_LBound))
    
    Dim StartingOffset As Long
    StartingOffset = OutputOffset
    
    Dim OutputBufferUB As Long
    OutputBufferUB = UBound(OutputBuffer)
    
    Do While InputCount > 0
        Dim b As Long
        b = InputBuffer(InputOffset)
        If CanProcessChar(b) Then
            If b <> vbBase64TerminatorChar Then
                If mTermCount > 0 Then _
                    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
                    
                b = Base64CharToBits(b)
                If b = corInvalidChar Then _
                    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
                    
                mBits = mBits Or b
            Else
                mTermCount = mTermCount + 1
                If mTermCount > 2 Then _
                    Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
                
            End If
            mBitCount = mBitCount + 6
            
            If mBitCount <> 24 Then
                mBits = mBits * &H40&
            Else
                Dim NewOffset As Long
                NewOffset = OutputOffset + (2 - mTermCount)
                If NewOffset > OutputBufferUB Then Call SmallBufferError("OutputBuffer")
                
                OutputBuffer(OutputOffset) = (mBits And &HFF0000) \ &H10000
                
                Select Case mTermCount
                    Case 0
                        OutputBuffer(OutputOffset + 2) = mBits And &HFF&
                        OutputBuffer(OutputOffset + 1) = (mBits And &HFF00&) \ &H100&

                    Case 1
                        OutputBuffer(OutputOffset + 1) = (mBits And &HFF00&) \ &H100&

                End Select
                
                OutputOffset = NewOffset + 1
                Call Reset
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
    Dim Result As Long
    Result = VerifyArrayRange(SAPtr(InputBuffer), InputOffset, InputCount)
    If Result <> NO_ERROR Then Call ThrowArrayRangeException(Result, "InputBuffer", InputOffset, "InputOffset", InputCount, "InputCount", False)

    Dim Ret() As Byte
    If (mBitCount \ 6) + InputCount < 4 Then
        Ret = Cor.NewBytes()
    ElseIf mTermCount > 0 Then
        ReDim Ret(0 To 1 - mTermCount)
        Call TransformBlock(InputBuffer, InputOffset, InputCount, Ret, 0)
    Else
        ReDim Ret(0 To InputCount)
        
        Dim i As Long
        i = TransformBlock(InputBuffer, InputOffset, InputCount, Ret, 0)
        ReDim Preserve Ret(0 To i - 1)
    End If
        
    Call Reset
    TransformFinalBlock = Ret
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal WhiteSpaces As FromBase64TransformMode)
    mIgnoreWhiteSpaces = (WhiteSpaces = IgnoreWhiteSpaces)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SmallBufferError(ByVal ParamName As String)
    Throw Cor.NewArgumentException(Environment.GetResourceString(Argument_SmallConversionBuffer), ParamName)
End Sub

Private Function CanProcessChar(ByVal Ch As Long) As Boolean
    If IsWhiteSpace(Ch) Then
        If Not mIgnoreWhiteSpaces Then
            Throw Cor.NewFormatException(Environment.GetResourceString(Format_InvalidBase64Character))
        End If
    Else
        CanProcessChar = True
    End If
End Function

Private Function IsWhiteSpace(ByVal Ch As Long) As Boolean
    Select Case Ch
        Case &H20, &H9, &HA, &HB, &HC, &HD, &H85, &HA0: IsWhiteSpace = True
    End Select
End Function

Private Sub Reset()
    mTermCount = 0
    mBits = 0
    mBitCount = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
     mIgnoreWhiteSpaces = DEF_IGNOREWHITESPACES
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICryptoTransform Interface
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

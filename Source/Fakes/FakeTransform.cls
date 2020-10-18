VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICryptoTransform

Public CanReuseTransform            As Boolean
Public CanTransformMultipleBlocks   As Boolean
Public InputBlockSize               As Long
Public OutputBlockSize              As Long
Public Enabled                      As Boolean
Public TransformBlockCount          As New Count
Public TransformFinalBlockCount     As New Count
Public EnforceBlockSizeMultiple     As Boolean


Public Sub Init(ByVal Enabled As Boolean, Optional CanReuse As Boolean, Optional CanTransformMultiple As Boolean, Optional BlockSize As Long)
    Me.Enabled = Enabled
    Me.CanReuseTransform = CanReuse
    Me.CanTransformMultipleBlocks = CanTransformMultiple
    Me.InputBlockSize = BlockSize
    Me.OutputBlockSize = BlockSize
End Sub

Private Sub Class_Initialize()
    Enabled = True
    InputBlockSize = 8
    OutputBlockSize = 8
End Sub

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
    Me.TransformBlockCount.Increment
    
    If EnforceBlockSizeMultiple Then
        If InputCount Mod InputBlockSize <> 0 Then
            Err.Raise 5, "TransformBlock", "InputCount must be exact multple of InputBlockSize"
        End If
    End If
    
    If Enabled And InputCount > 0 Then
        CopyMemory OutputBuffer(OutputOffset), InputBuffer(InputOffset), InputCount
        ICryptoTransform_TransformBlock = InputCount
    End If
End Function

Private Function ICryptoTransform_TransformFinalBlock(InputBuffer() As Byte, ByVal InputOffset As Long, ByVal InputCount As Long) As Byte()
    Me.TransformFinalBlockCount.Increment
    
    If Enabled And InputCount > 0 Then
        Dim Result() As Byte
        ReDim Result(0 To InputCount - 1)
        CopyMemory Result(0), InputBuffer(InputOffset), InputCount
        ICryptoTransform_TransformFinalBlock = Result
    Else
        ICryptoTransform_TransformFinalBlock = NewBytes
    End If
End Function

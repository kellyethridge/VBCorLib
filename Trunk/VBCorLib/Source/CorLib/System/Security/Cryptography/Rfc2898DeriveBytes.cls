VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rfc2898DeriveBytes"
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
' Module: Rfc2898DeriveBytes
'

''
' Provides a methods to derive Byte array keys from string passwords.
'
' @remarks This class implements the RFC2898 standard for generating keys from string
' or byte array passwords utilizing the SHA-1 hash algorithm.
' <p>With this class, a key of any length can be generated from a string or byte array password.</p>
'
' @see Constructors
'
Option Explicit
Implements IObject

Private Const BLOCK_SIZE As Long = 20   ' the length of the computed hash from HMACSHA1


Private mPRF                As New HMACSHA1
Private mIterations         As Long
Private mSalt()             As Byte

' State for between calls to GetBytes
Private mBlockCount         As Long
Private mBuffer(0 To 19)    As Byte
Private mBytesBuffered      As Long



''
' Returns the number of iterations used when computing the next set of bytes.
'
' @Return The number of iterations to be used.
'
Public Property Get IterationCount() As Long
    IterationCount = mIterations
End Property

''
' Sets the number of iterations to be used when computing the next set of bytes.
'
' @param RHS The number of iterations to be used.
' @remarks The value must be greater than 0.
'
Public Property Let IterationCount(ByVal RHS As Long)
    If RHS < 1 Then _
        Throw Cor.NewArgumentException("Cannot set iteration count less than 1.", "IterationCount")
    
    mIterations = RHS
End Property

''
' Returns any Salt being used during key generation.
'
' @return The Salt being used during key generation.
'
Public Property Get Salt() As Byte()
    Salt = mSalt
End Property

''
' Sets the Salt to be used during key generation.
'
' @param RHS The Salt to be used.
' @remarks The length of the array cannot be less than 8 bytes.
'
Public Property Let Salt(ByRef RHS() As Byte)
    If CorArray.IsNull(RHS) Then _
        Throw Cor.NewArgumentNullException("Salt")
    
    Dim Size As Long
    Size = CorArray.GetLength(RHS)
    
    If Size < 8 Then _
        Throw Cor.NewArgumentException("Cannot set Salt less than 8 bytes in length.", "Salt")
    
    ReDim mSalt(0 To Size - 1)
    Call Buffer.BlockCopy(RHS, 0, mSalt, 0, Size)
End Property

''
' Returns the next set of bytes to be generated for a specific string password.
'
' @param cb The number of bytes to be generated.
'
Public Function GetBytes(ByVal cb As Long) As Byte()
    Dim Blocks As Long
    Blocks = MathExt.Ceiling(cb / BLOCK_SIZE)
    
    Dim Bytes() As Byte
    ReDim Bytes(0 To (Blocks * BLOCK_SIZE) + mBytesBuffered - 1)
    
    If mBytesBuffered > 0 Then Call CopyMemory(Bytes(0), mBuffer(0), mBytesBuffered)
    
    Dim i As Long
    For i = 1 To Blocks
        Dim BlockBytes() As Byte
        BlockBytes = NextBlockBytes
        Call CopyMemory(Bytes((i - 1) * BLOCK_SIZE + mBytesBuffered), BlockBytes(0), BLOCK_SIZE)
    Next i
    
    mBytesBuffered = (Blocks * BLOCK_SIZE) - cb
    If mBytesBuffered > 0 Then Call CopyMemory(mBuffer(0), Bytes(cb), mBytesBuffered)
    
    ReDim Preserve Bytes(0 To cb - 1)
    GetBytes = Bytes
End Function

''
' Resets the key generator.
'
Public Sub Reset()
    mPRF.Initialize
    mBlockCount = 0
    mBytesBuffered = 0
End Sub

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
Friend Sub Init(ByRef Password As Variant, ByRef Salt As Variant, ByVal Iterations As Long)
    Me.IterationCount = Iterations
    
    Select Case VarType(Password)
        Case vbString:      mPRF.Key = Encoding.UTF8.GetBytes(Password)
        Case vbByteArray:   mPRF.Key = Password
        Case Else
            Throw Cor.NewArgumentException("Password must be a byte array or string", "Password")
    End Select

    Select Case VarType(Salt)
        Case vbByteArray:               Me.Salt = Salt
        Case vbLong, vbInteger, vbByte: Me.Salt = CryptoHelper.GetRandomBytes(Salt)
        Case Else
            Throw Cor.NewArgumentException("Salt must be a byte array or a number.", "Salt")
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextBlockBytes() As Byte()
    mBlockCount = mBlockCount + 1
    
    Call mPRF.TransformBlock(mSalt, 0, CorArray.GetLength(mSalt), mSalt, 0)
    
    Dim Result() As Byte
    Result = mPRF.ComputeHash(AsQByteArr(SwapEndian(mBlockCount)).Bytes)
    
    Dim sum() As Byte
    sum = Result
    
    Dim i As Long
    For i = 2 To mIterations
        Result = mPRF.ComputeHash(Result)
        
        Dim j As Long
        For j = 0 To BLOCK_SIZE - 1
            sum(j) = sum(j) Xor Result(j)
        Next j
    Next i
    
    NextBlockBytes = sum
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mIterations = 1000
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

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
' Implements password-based key derivation functionality, PBKDF2, by using a pseudo-random number generator based on <a href="HMACSHA1.html">HMACSHA1</a>.
'
' @see Constructors
' @include "..\DocIncludes\Rfc2898DeriveBytes.html"
Option Explicit
Implements IObject

Private Const BlockSize As Long = 20   ' the length of the computed hash from HMACSHA1

Private mHasher             As HMACSHA1
Private mIterations         As Long
Private mSalt()             As Byte

' State for between calls to GetBytes
Private mBlockCount         As Long
Private mBuffer(0 To 19)    As Byte
Private mBytesBuffered      As Long


''
' Gets the number of iterations for the operation.
'
' @Return The number of iterations for the operation.
' @include "..\DocIncludes\Rfc2898DeriveBytes.Get.IterationCount.html"
Public Property Get IterationCount() As Long
    IterationCount = mIterations
End Property

''
' Sets the number of iterations for the operation.
'
' @param Value The number of iterations for the operation.
' @include "..\DocIncludes\Rfc2898DeriveBytes.Let.IterationCount.html"
Public Property Let IterationCount(ByVal Value As Long)
    If Value < 1 Then _
        Error.ArgumentOutOfRange "Value", ArgumentOutOfRange_NeedPosNum
    
    mIterations = Value
    Reset
End Property

''
' Gets the key salt value for the operation.
'
' @return The key salt value for the operation.
' @remarks Salt, a random set of bytes, is used to make unauthorized decrypting of a message more difficult. A dictionary attack
' is an attack in which the attacker attempts to decrypt an encrypted message by comparing the encrypted value with previously
' computed encrypted values for the most likely keys. This attack is made much more difficult by the introduction of salt, or
' random bytes, at the end of the password before the key derivation.
Public Property Get Salt() As Byte()
    Salt = mSalt
End Property

''
' Sets the key salt value for the operation.
'
' @param Value The key salt value for the operation.
' @remarks The length of the array cannot be less than 8 bytes.
' @remarks Salt, a random set of bytes, is used to make unauthorized decrypting of a message more difficult. A dictionary attack
' is an attack in which the attacker attempts to decrypt an encrypted message by comparing the encrypted value with previously
' computed encrypted values for the most likely keys. This attack is made much more difficult by the introduction of salt, or
' random bytes, at the end of the password before the key derivation.
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The specified salt size is smaller than 8 bytes.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>The salt is an uninitialized array.</td>
'    </tr>
' </table>
Public Property Let Salt(ByRef Value() As Byte)
    ValidateArray Value, NameOfValue
    
    If Len1D(Value) < 8 Then _
        Error.Argument Cryptography_PasswordDerivedBytes_FewBytesSalt, "Value"
    
    mSalt = Value
    Reset
End Property

''
' Returns the pseudo-random key for this object.
'
' @param cb The number of pseudo-random key bytes to generate.
' @return A byte array filled with pseudo-random key bytes.
' @include "..\DocIncludes\Rfc2898DeriveBytes.GetBytes.html"
Public Function GetBytes(ByVal cb As Long) As Byte()
    Dim Blocks          As Long
    Dim Bytes()         As Byte
    Dim BlockBytes()    As Byte
    Dim i               As Long
        
    If cb < 0 Then _
        Error.ArgumentOutOfRange "cb", ArgumentOutOfRange_NeedNonNegNum
        
    Blocks = CorMath.Ceiling(cb / BlockSize)
    ReDim Bytes(0 To (Blocks * BlockSize) + mBytesBuffered - 1)
    
    If mBytesBuffered > 0 Then
        CopyMemory Bytes(0), mBuffer(0), mBytesBuffered
    End If
    
    For i = 1 To Blocks
        BlockBytes = NextBlockBytes
        CopyMemory Bytes((i - 1) * BlockSize + mBytesBuffered), BlockBytes(0), BlockSize
    Next i
    
    mBytesBuffered = (Blocks * BlockSize) - cb
    
    If mBytesBuffered > 0 Then
        CopyMemory mBuffer(0), Bytes(cb), mBytesBuffered
    End If
    
    ReDim Preserve Bytes(0 To cb - 1)
    GetBytes = Bytes
End Function

''
' Resets the state of the operation.
'
' @remarks This method is automatically called if the salt or iteration count is modified.
Public Sub Reset()
    mHasher.Initialize
    mBlockCount = 0
    mBytesBuffered = 0
    CorArray.Clear mBuffer, 0, 20
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
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Password As Variant, ByRef Salt As Variant, ByVal Iterations As Long)
    InitHasher Password
    InitSalt Salt
    IterationCount = Iterations
End Sub

Private Sub InitHasher(ByRef Password As Variant)
    Dim Key() As Byte
    
    Select Case VarType(Password)
        Case vbString
            Key = Encoding.UTF8.GetBytes(Password)
        Case vbByteArray
            Key = Password
        Case Else
            Error.Argument Argument_ByteArrayOrStringRequired, "Password"
    End Select
    
    Set mHasher = Cor.NewHMACSHA1(Key)
End Sub

Private Sub InitSalt(ByRef Salt As Variant)
    Select Case VarType(Salt)
        Case vbByteArray
            Me.Salt = Salt
        Case vbLong, vbInteger, vbByte
            If Salt < 0 Then _
                Error.ArgumentOutOfRange "Salt", ArgumentOutOfRange_NeedNonNegNum
                
            Me.Salt = CryptoHelper.GetRandomBytes(Salt)
        Case Else
            Error.Argument Argument_ByteArrayOrNumberRequired, "Salt"
    End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function NextBlockBytes() As Byte()
    mBlockCount = mBlockCount + 1
    
    mHasher.TransformBlock mSalt, LBound(mSalt), CorArray.Length(mSalt), mSalt, LBound(mSalt)
    
    Dim Result() As Byte
    Result = mHasher.ComputeHash(AsQByteArr(SwapEndian(mBlockCount)).Bytes)
    
    Dim sum() As Byte
    sum = Result
    
    Dim i As Long
    For i = 2 To mIterations
        Result = mHasher.ComputeHash(Result)
        
        Dim j As Long
        For j = 0 To BlockSize - 1
            sum(j) = sum(j) Xor Result(j)
        Next j
    Next i
    
    NextBlockBytes = sum
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

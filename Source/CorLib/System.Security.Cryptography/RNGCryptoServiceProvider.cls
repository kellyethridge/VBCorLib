VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RNGCryptoServiceProvider"
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
' Module: RNGCryptoServiceProvider
'

''
' Provides random number generation that is a bit more random than the regular Visual Basic method
' by using a Cryptographic Service Provider.
'
'@Folder("CorLib.System.Security.Cryptography")
Option Explicit
Implements IObject
Implements RandomNumberGenerator

Private mProvider As Long


''
' Fills an array of bytes with random values.
'
' @param Data The byte array to be filled.
'
Public Sub GetBytes(ByRef Data() As Byte)
    VerifyProvider
    
    If CorArray.IsNull(Data) Then _
        Error.ArgumentNull "Data", ArgumentNull_Array
        
    If CryptGenRandom(mProvider, CorArray.Length(Data), Data(LBound(Data))) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetMessage(Err.LastDllError))
End Sub

''
' Fills an array of bytes with random values, with no zeros.
'
' @param Data The byte array to be filled.
'
Public Sub GetNonZeroBytes(ByRef Data() As Byte)
    Dim BigData() As Byte
    
    ' We double the size of the array and fill it with
    ' random values, banking on the odds that half of
    ' it will not be zeros.
    ReDim BigData(CorArray.Length(Data) * 2 - 1)
    GetBytes BigData
    
    Dim ubData  As Long
    Dim j       As Long
    Dim i       As Long
    
    j = LBound(Data)
    ubData = UBound(Data)
    
    For i = 0 To UBound(BigData)
        If BigData(i) <> 0 Then
            Data(j) = BigData(i)
            j = j + 1
            If j > ubData Then
                Exit Sub
            End If
        End If
    Next i
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
Friend Sub Init(ByVal CspParams As CspParameters)
    If CspParams Is Nothing Then _
        Error.ArgumentNull "CspParams"
    
    With CspParams
        mProvider = CryptoAPI.AcquireContext(.KeyContainerName, .ProviderName, .ProviderType, .Flags)
    End With
End Sub

Private Sub Class_Terminate()
    If mProvider <> 0 Then
        CryptoAPI.ReleaseContext mProvider, True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyProvider()
    If mProvider = vbNullPtr Then
        Init Cor.NewCspParameters(PROV_RSA_AES)
    End If
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
'   RandomNumberGenerator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RandomNumberGenerator_Equals(Value As Variant) As Boolean
    RandomNumberGenerator_Equals = Equals(Value)
End Function

Private Sub RandomNumberGenerator_GetBytes(Data() As Byte)
    GetBytes Data
End Sub

Private Function RandomNumberGenerator_GetHashCode() As Long
    RandomNumberGenerator_GetHashCode = GetHashCode
End Function

Private Sub RandomNumberGenerator_GetNonZeroBytes(Data() As Byte)
    GetNonZeroBytes Data
End Sub

Private Function RandomNumberGenerator_ToString() As String
    RandomNumberGenerator_ToString = ToString
End Function

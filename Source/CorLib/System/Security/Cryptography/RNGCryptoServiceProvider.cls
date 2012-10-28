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
'    Module: RNGCryptoServiceProvider
'

''
' Provides random number generation that is a bit more random than the regular Visual Basic method
' by using a Cryptographic Service Provider.
'
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
    Call VerifyProvider
    
    If cArray.IsNull(Data) Then _
        Throw Cor.NewArgumentNullException(Environment.GetResourceString(ArgumentNull_Array), "Data")
    
    If CryptGenRandom(mProvider, cArray.GetLength(Data), Data(LBound(Data))) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetErrorMessage(Err.LastDllError))
    
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
    ReDim BigData(cArray.GetLength(Data) * 2 - 1)
    Call GetBytes(BigData)
    
    Dim j As Long
    j = LBound(Data)
    
    Dim ub As Long
    ub = UBound(Data)
    
    Dim i As Long
    For i = 0 To UBound(BigData)
        If BigData(i) <> 0 Then
            Data(j) = BigData(i)
            j = j + 1
            If j > ub Then Exit Sub
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
    Equals = Object.Equals(Me, Value)
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
    GetHashCode = ObjPtr(CUnk(Me))
End Function

''
' Returns a string representation of this object instance.
' The default method simply returns the application name
' and class name in which this class resides.
'
' A Person class may return the person's name instead.
'
Public Function ToString() As String
    ToString = Object.ToString(Me, App)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal CspParams As CspParameters)
    If CspParams Is Nothing Then _
        Throw Cor.NewArgumentNullException("CspParams cannot be Nothing.", "CspParams")
    
    With CspParams
        mProvider = CryptoAPI.AcquireContext(.KeyContainerName, .ProviderName, .ProviderType, .Flags)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyProvider()
    If mProvider = 0 Then Call Init(Cor.NewCspParameters(PROV_RSA_FULL))
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    If mProvider <> 0 Then Call CryptoAPI.ReleaseContext(mProvider, True)
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   RandomNumberGenerator Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RandomNumberGenerator_Equals(Value As Variant) As Boolean
    RandomNumberGenerator_Equals = Equals(Value)
End Function

Private Sub RandomNumberGenerator_GetBytes(Data() As Byte)
    Call GetBytes(Data)
End Sub

Private Function RandomNumberGenerator_GetHashCode() As Long
    RandomNumberGenerator_GetHashCode = GetHashCode
End Function

Private Sub RandomNumberGenerator_GetNonZeroBytes(Data() As Byte)
    Call GetNonZeroBytes(Data)
End Sub

Private Function RandomNumberGenerator_ToString() As String
    RandomNumberGenerator_ToString = ToString
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CspKeyContainerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: CspKeyContainerInfo
'

''
' Represents information about a Crypto Service Provider's Key Container.
'
' @remarks To create an instance of <b>CspKeyContainerInfo</b> use the <a href="Constructors.NewCspKeyContainerInfo.html">NewCspKeyContainerInfo</a> constructor.
' @see Constructors
Option Explicit
Implements IObject

Private mProviderName           As String
Private mProviderType           As Long
Private mKeyContainerName       As String
Private mUniqueKeyContainerName As String
Private mKeyNumber              As KeyNumber
Private mMachineKeyStore        As Boolean
Private mExportable             As Boolean
Private mAccessible             As Boolean
Private mRandomlyGenerated      As Boolean
Private mRemovable              As Boolean
Private mIsLoaded               As Boolean



''
' Returns if a key in a key container is accessible.
'
' @return Returns True if the key is accessible, False otherwise.
'
Public Property Get Accessible() As Boolean
    Accessible = mAccessible
End Property

Friend Property Let Accessible(ByVal Value As Boolean)
    mAccessible = Value
End Property

''
' Returns if a key in a key container can be exported out of the container.
'
' @return Returns True if the key can be exported, False otherwise.
'
Public Property Get Exportable() As Boolean
    VerifyLoaded
    Exportable = mExportable
End Property

''
' Returns the name of the key container.
'
' @return The name of the key container.
'
Public Property Get KeyContainerName() As String
    KeyContainerName = mKeyContainerName
End Property

''
' Returns the key number, representing the type of key.
'
' @return Returns 1 for Exchange, 2 for Signature.
'
Public Property Get KeyNumber() As KeyNumber
    KeyNumber = mKeyNumber
End Property

''
' Returns if the key is stored in a machine storage container.
'
' @Return Returns True if the key is in a machine set, False otherwise.
'
Public Property Get MachineKeyStore() As Boolean
    MachineKeyStore = mMachineKeyStore
End Property

''
' Returns if the key is protected.
'
' @return Returns True if protected, False otherwise.
'
Public Property Get Protected() As Boolean
End Property

''
' Returns the name of the cryptography service provider.
'
' @return A string containing the name of the provider.
'
Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property

''
' Returns the type of cryptography service provider.
'
' @return Returns the type of the provider.
'
Public Property Get ProviderType() As Long
    ProviderType = mProviderType
End Property

''
' Returns if the key container was randomly generated or imported.
'
' @return Returns True if the key container was randomly generated, False otherwise.
'
Public Property Get RandomlyGenerated() As Boolean
    RandomlyGenerated = mRandomlyGenerated
End Property

Friend Property Let RandomlyGenerated(ByVal Value As Boolean)
    mRandomlyGenerated = Value
End Property

''
' Returns if the key can be removed from the key container.
'
' @return Returns True if the key can be removed, False otherwise.
'
Public Property Get Removable() As Boolean
    Removable = mRemovable
End Property

''
' Returns a unique key container name.
'
' @return A unique key container name.
'
Public Property Get UniqueKeyContainerName() As String
    VerifyLoaded
    UniqueKeyContainerName = mUniqueKeyContainerName
End Property

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
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
Friend Sub Init(ByVal Parameters As CspParameters)
    If Parameters Is Nothing Then _
        Error.ArgumentNull "Parameters"
        
    mAccessible = True
    
    With Parameters
        mProviderName = .ProviderName
        mProviderType = .ProviderType
        mKeyContainerName = .KeyContainerName
        
        If .KeyNumber = -1 Then
            If .ProviderType = PROV_RSA_FULL Or .ProviderType = PROV_RSA_AES Then
                mKeyNumber = CorLib.KeyNumber.Exchange
            ElseIf .ProviderType = PROV_DSS_DH Then
                mKeyNumber = CorLib.KeyNumber.Signature
            Else
                mKeyNumber = .KeyNumber
            End If
        Else
            mKeyNumber = .KeyNumber
        End If
        
        mMachineKeyStore = CBool(.Flags And UseMachineKeyStore)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyLoaded()
    If mIsLoaded Then
        Exit Sub
    End If
    
    Dim Provider As Long
    Provider = CryptoAPI.AcquireContext(mKeyContainerName, , mProviderType)
    mUniqueKeyContainerName = CryptoAPI.GetProvParamString(Provider, PP_UNIQUE_CONTAINER)
    
    Dim Key As Long
    If CryptGetUserKey(Provider, mKeyNumber, Key) = BOOL_FALSE Then
        CryptoAPI.ReleaseContext Provider
        Throw Cor.NewCryptographicException(GetMessage(Err.LastDllError))
    End If
    
    Dim Permissions As Long
    Permissions = CryptoAPI.GetKeyParamLong(Key, KP_PERMISSIONS)
    mExportable = (Permissions And CRYPT_EXPORT)
    
    If Len(mProviderName) = 0 Then
        mProviderName = CryptoAPI.GetProvParamString(Provider, PP_NAME)
    End If
    
    CryptoAPI.ReleaseContext Provider, True
    mIsLoaded = True
End Sub


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

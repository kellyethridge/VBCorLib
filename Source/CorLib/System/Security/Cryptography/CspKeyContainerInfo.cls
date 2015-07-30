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
'    Module: CspKeyContainerInfo
'

''
' Represents information about a Crypto Service Provider's Key Container.
'
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
Private mProtected              As Boolean
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

''
' Returns if a key in a key container can be exported out of the container.
'
' @return Returns True if the key can be exported, False otherwise.
'
Public Property Get Exportable() As Boolean
    Call VerifyLoaded
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
    Call VerifyLoaded
    Protected = mProtected
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
    Call VerifyLoaded
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Parameters As CspParameters)
    mAccessible = True
    
    With Parameters
        mProviderName = .ProviderName
        mProviderType = .ProviderType
        mKeyContainerName = .KeyContainerName
        mKeyNumber = IIf(.KeyNumber = -1, CorLib.KeyNumber.Exchange, .KeyNumber)
        mMachineKeyStore = CBool(.Flags And UseMachineKeyStore)
    End With
End Sub

Friend Property Let RandomlyGenerated(ByVal RHS As Boolean)
    mRandomlyGenerated = RHS
End Property

Friend Property Let Accessible(ByVal RHS As Boolean)
    mAccessible = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyLoaded()
    If mIsLoaded Then Exit Sub
    
    On Error GoTo 0
    
    Dim prov As Long
    prov = CryptoAPI.AcquireContext(mKeyContainerName, mProviderName, mProviderType)
    
    mUniqueKeyContainerName = CryptoAPI.GetProvParamString(prov, PP_UNIQUE_CONTAINER)
    
    Dim Key As Long
    If CryptGetUserKey(prov, mKeyNumber, Key) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetMessage(Err.LastDllError))
    
    Dim Permissions As Long
    Permissions = CryptoAPI.GetKeyParamLong(Key, KP_PERMISSIONS)
    mExportable = (Permissions And CRYPT_EXPORT)
    
    If Len(mProviderName) = 0 Then
        mProviderName = CryptoAPI.GetProvParamString(prov, PP_NAME)
    End If
    
errTrap:
    Call CryptoAPI.ReleaseContext(prov, True)
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

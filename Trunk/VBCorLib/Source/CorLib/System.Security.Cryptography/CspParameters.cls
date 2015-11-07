VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CspParameters"
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
' Module: CspParameters
'

''
' Contains parameter values that can be passed on to cryptography providers.
'
Option Explicit
Implements IObject

''
' A set of flags that can be used to modify the behavior of cryptographic providers.
'
' @param NoFlags No flags are to be set.
' @param NoPrompt Prevents the service provider from display any UI for this context.
' @param UseArchivableKey Allows a key to be exported for archival or recovery.
' @param UseDefaultKeyContainer Use the key information from the default container.
' @param UseExistingKey Use the key information from the current key.
' @param UseMachineKeyStore Use the key information from the computer's key store.
' @param UseNonExportableKey Use key information that cannot be exported.
' @param UseUserProtectedKey Notify the user through a dialog box or another method
' when certain actions are attempting to use a key. This flag is not compatible with the NoPrompt flag.
'
Public Enum CspProviderFlags
    NoFlags = 0
    NoPrompt = 64
    UseArchivableKey = 16
    UseDefaultKeyContainer = 2
    UseExistingKey = 8
    UseMachineKeyStore = 1
    UseNonExportableKey = 4
    UseUserProtectedKey = 32
End Enum

Private Const DefaultKeyNumber      As Long = -1

Private mKeyContainerName   As String
Private mKeyNumber          As Long
Private mProviderName       As String
Private mProviderType       As Long
Private mFlags              As CspProviderFlags


''
' Returns the key container name.
'
' @return The name of the key container holding a key.
' @remarks A key can be kept within a container for later retrieval. This is the name of that container.
'
Public Property Get KeyContainerName() As String
    KeyContainerName = mKeyContainerName
End Property

''
' Sets the key container name.
'
' @param Value The key container name that will hold a key.
' @remarks A key can be kept within a container for later retrieval. This is the name of that container.
'
Public Property Let KeyContainerName(ByVal Value As String)
    mKeyContainerName = Value
End Property

''
' Returns if an Asymmetric algorithm key is a Signature or Exchange key.
'
' @return Returns 1 for Exchange, or 2 for Signature.
' @remarks An Exchange key is a public/private key pair use for Asymmetric encryption. A Signature is a
' key pair used to digitally sign a message or file.
'
Public Property Get KeyNumber() As Long
    KeyNumber = mKeyNumber
End Property

''
' Sets if an Asymmetric algorithm key is a Signature or Exchange key.
'
' @param Value Set to 1 for Exchange, or 2 for Signature.
' @remarks An Exchange key is a public/private key pair use for Asymmetric encryption. A Signature is a
' key pair used to digitally sign a message or file.
'
Public Property Let KeyNumber(ByVal Value As Long)
    mKeyNumber = Value
End Property

''
' Returns the Provider name of the crypto service provider.
'
' @return Returns the name of the provider.
'
Public Property Get ProviderName() As String
    ProviderName = mProviderName
End Property

''
' Sets the name of the provider to create when acquiring a crypto service provider.
'
' @param Value The provider name to acquire.
'
Public Property Let ProviderName(ByVal Value As String)
    mProviderName = Value
End Property

''
' Returns the provider type code.
'
' @return The code of the provider.
' @remarks Some provider type codes are:
' <pre>
' PROV_RSA_FULL         = 1
' PROV_RSA_SIG          = 2
' PROV_DSS              = 3
' PROV_FORTEZZA         = 4
' PROV_MS_EXCHANGE      = 5
' PROV_SSL              = 6
' PROV_RSA_SCHANNEL     = 12
' PROV_DSS_DH           = 13
' PROV_EC_ECDSA_SIG     = 14
' PROV_EC_ECNRA_SIG     = 15
' PROV_EC_ECDSA_FULL    = 16
' PROV_EC_ECNRA_FULL    = 17
' PROV_DH_SCHANNEL      = 18
' PROV_SPYRUS_LYNKS     = 20
' PROV_RNG              = 21
' PROV_INTEL_SEC        = 22
' PROV_REPLACE_OWF      = 23
' PROV_RSA_AES          = 24
'</pre>
'
Public Property Get ProviderType() As Long
    ProviderType = mProviderType
End Property

''
' Sets the provider type code.
'
' @param Value The provider type code.
' @remarks Some provider type codes are:
' <pre>
' PROV_RSA_FULL         = 1
' PROV_RSA_SIG          = 2
' PROV_DSS              = 3
' PROV_FORTEZZA         = 4
' PROV_MS_EXCHANGE      = 5
' PROV_SSL              = 6
' PROV_RSA_SCHANNEL     = 12
' PROV_DSS_DH           = 13
' PROV_EC_ECDSA_SIG     = 14
' PROV_EC_ECNRA_SIG     = 15
' PROV_EC_ECDSA_FULL    = 16
' PROV_EC_ECNRA_FULL    = 17
' PROV_DH_SCHANNEL      = 18
' PROV_SPYRUS_LYNKS     = 20
' PROV_RNG              = 21
' PROV_INTEL_SEC        = 22
' PROV_REPLACE_OWF      = 23
' PROV_RSA_AES          = 24
'</pre>
'
Public Property Let ProviderType(ByVal Value As Long)
    mProviderType = Value
End Property

''
' Returns flags used to modify the behavior of cryptographic providers.
'
' @return The flags set for cryptographic providers.
'
Public Property Get Flags() As CspProviderFlags
    Flags = mFlags
End Property

''
' Sets flags used to modify the behavior of cryptographic providers.
'
' @param Value The flags set for cryptographic providers.
'
Public Property Let Flags(ByVal Value As CspProviderFlags)
    If (Value And &HFFFFFF80) <> 0 Then _
        Throw Cor.NewArgumentException("Invalid CspParameter flag value.", "Value")
    
    mFlags = Value
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
Friend Sub Init(ByVal TypeIn As Long, ByVal ProviderNameIn As String, ByVal ContainerNameIn As String)
    mProviderType = TypeIn
    mProviderName = ProviderNameIn
    mKeyContainerName = ContainerNameIn
End Sub

Private Sub Class_Initialize()
    mFlags = NoFlags
    mKeyNumber = DefaultKeyNumber
    mProviderType = PROV_RSA_AES
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

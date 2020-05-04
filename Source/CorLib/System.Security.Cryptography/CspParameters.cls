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
' @param UseUserProtectedKey Notify the user through a dialog box or another method when certain actions are attempting to use a key. This flag is not compatible with the NoPrompt flag.
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

Private Const DefaultKeyNumber As Long = -1

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
' Gets the provider type code for <a href="CspParameters.html">CspParameters</a>.
'
' @return The numeric code of the provider.
' @remarks Use the <b>ProviderType</b> property to specify a provider type by passing a numeric value that represents that provider.
' The numeric values that represent the default provider types are defined in the WinCrypt.h header file:
' <ul>
'   <li>To specify a provider compatible with the <a href="RSA.html">RSA</a> class, pass a value of 1.</li>
'   <li>To specify a provider compatible with the <a href="DSA.html">DSA</a> class, pass a value of 13.</li>
' </ul>
' The following table lists the provider types and their corresponding numeric values defined in the WinCrypt.h header file.
' <table style="width: auto">
'   <tr>
'       <th>Provider Type</th><th>Numeric Value</th>
'   </tr>
'   <tr>
'        <td>PROV_RSA_RULL</td><td>1</td>
'   </tr>
'    <tr>
'        <td>PROV_RSA_SIG</td><td>2</td>
'    </tr>
'    <tr>
'        <td>PROV_DSS</td><td>3</td>
'    </tr>
'    <tr>
'        <td>PROV_FORTEZZA</td><td>4</td>
'    </tr>
'    <tr>
'        <td>PROV_MS_EXCHANGE</td><td>5</td>
'    </tr>
'    <tr>
'        <td>PROV_SSL</td><td>6</td>
'    </tr>
'    <tr>
'        <td>PROV_RSA_SCHANNEL</td><td>12</td>
'    </tr>
'    <tr>
'        <td>PROV_DSS_DH</td><td>13</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECDSA_SIG</td><td>14</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECNRA_SIG</td><td>15</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECDSA_FULL</td><td>16</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECNRA_FULL</td><td>17</td>
'    </tr>
'    <tr>
'        <td>PROV_DH_SCHANNEL</td><td>18</td>
'    </tr>
'    <tr>
'        <td>PROV_SPYRUS_LYNKS</td><td>20</td>
'    </tr>
'    <tr>
'        <td>PROV_RNG</td><td>21</td>
'    </tr>
'    <tr>
'        <td>PROV_INTEL_SEC</td><td>22</td>
'    </tr>
'    <tr>
'        <td>PROV_REPLACE_OWF</td><td>23</td>
'    </tr>
'    <tr>
'        <td>PROV_RSA_AES</td><td>24</td>
'    </tr>
' </table>
' For more information about the default provider types and their behaviors, see the Microsoft Cryptography API (CAPI) documentation in the MSDN library.
Public Property Get ProviderType() As Long
    ProviderType = mProviderType
End Property

''
' Sets the provider type code for <a href="CspParameters.html">CspParameters</a>.
'
' @param Value The numeric code of the provider.
' @remarks Use the <b>ProviderType</b> property to specify a provider type by passing a numeric value that represents that provider.
' The numeric values that represent the default provider types are defined in the WinCrypt.h header file:
' <ul>
'   <li>To specify a provider compatible with the <a href="RSA.html">RSA</a> class, pass a value of 1.</li>
'   <li>To specify a provider compatible with the <a href="DSA.html">DSA</a> class, pass a value of 13.</li>
' </ul>
' The following table lists the provider types and their corresponding numeric values defined in the WinCrypt.h header file.
' <table style="width: auto">
'   <tr>
'       <th>Provider Type</th><th>Numeric Value</th>
'   </tr>
'   <tr>
'        <td>PROV_RSA_RULL</td><td>1</td>
'   </tr>
'    <tr>
'        <td>PROV_RSA_SIG</td><td>2</td>
'    </tr>
'    <tr>
'        <td>PROV_DSS</td><td>3</td>
'    </tr>
'    <tr>
'        <td>PROV_FORTEZZA</td><td>4</td>
'    </tr>
'    <tr>
'        <td>PROV_MS_EXCHANGE</td><td>5</td>
'    </tr>
'    <tr>
'        <td>PROV_SSL</td><td>6</td>
'    </tr>
'    <tr>
'        <td>PROV_RSA_SCHANNEL</td><td>12</td>
'    </tr>
'    <tr>
'        <td>PROV_DSS_DH</td><td>13</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECDSA_SIG</td><td>14</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECNRA_SIG</td><td>15</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECDSA_FULL</td><td>16</td>
'    </tr>
'    <tr>
'        <td>PROV_EC_ECNRA_FULL</td><td>17</td>
'    </tr>
'    <tr>
'        <td>PROV_DH_SCHANNEL</td><td>18</td>
'    </tr>
'    <tr>
'        <td>PROV_SPYRUS_LYNKS</td><td>20</td>
'    </tr>
'    <tr>
'        <td>PROV_RNG</td><td>21</td>
'    </tr>
'    <tr>
'        <td>PROV_INTEL_SEC</td><td>22</td>
'    </tr>
'    <tr>
'        <td>PROV_REPLACE_OWF</td><td>23</td>
'    </tr>
'    <tr>
'        <td>PROV_RSA_AES</td><td>24</td>
'    </tr>
' </table>
' For more information about the default provider types and their behaviors, see the Microsoft Cryptography API (CAPI) documentation in the MSDN library.
Public Property Let ProviderType(ByVal Value As Long)
    mProviderType = Value
End Property

''
' Gets the flags for <a href="CspParameters.html">CspParameters</a> that modify the behavior of the cryptographic service provider (CSP).
'
' @return An enumeration value, or a bitwise combination of enumeration values.
Public Property Get Flags() As CspProviderFlags
    Flags = mFlags
End Property

''
' Sets the flags for <a href="CspParameters.html">CspParameters</a> that modify the behavior of the cryptographic service provider (CSP).
'
' @param Value An enumeration value, or a bitwise combination of enumeration values.
' @exceptions
'' <table style="width: auto">
 '   <tr>
 '       <th>Exception</th><th>Condition</th>
 '   </tr>
 '   <tr>
 '        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Value</i> is not a valid enumeration value.</td>
 '   </tr>
 ' </table>
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

Friend Function Clone() As CspParameters
    Set Clone = New CspParameters
    Clone.InitClone Me
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    mFlags = NoFlags
    mKeyNumber = DefaultKeyNumber
    mProviderType = PROV_RSA_AES
End Sub

Friend Sub Init(ByVal TypeIn As Long, ByRef ProviderNameIn As String, ByRef ContainerNameIn As String)
    mProviderType = TypeIn
    mProviderName = ProviderNameIn
    mKeyContainerName = ContainerNameIn
End Sub

Friend Sub InitClone(ByVal Params As CspParameters)
    Init Params.ProviderType, Params.ProviderName, Params.KeyContainerName
    mKeyNumber = Params.KeyNumber
    mFlags = Params.Flags
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

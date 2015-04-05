VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSACryptoServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' Module: DSACryptoServiceProvider
'

''
' Provides a wrapper to access a DSA crypto service provider.
'
' @remarks This class is used to create signatures for data and hashes and provide
' verification of those signatures generated.
'
Option Explicit
Implements IObject
Implements AsymmetricAlgorithm
Implements DSA
Implements ICspAsymmetricAlgorithm

Private Const DEF_KEYNUMBER     As Long = AT_SIGNATURE
Private Const DEF_PROVIDERTYPE  As Long = PROV_DSS_DH

Private mProvider           As Long
Private mKey                As Long
Private mKeySize            As Long
Private mProviderName       As String
Private mProviderType       As Long
Private mKeyContainerName   As String
Private mKeyNumber          As Long
Private mPersistKeyInCsp    As Boolean
Private mFlags              As CspProviderFlags
Private mLegalKeySizes()    As KeySizes
Private mRandomlyGenerated  As Boolean
Private mPublicOnly         As Boolean
Private mIsDisposed         As Boolean




''
' Returns a CspKeyContainerInfo object containing information about this DSACryptoServiceProvider instance.
'
' @return A CspKeyContainerInfo object.
'
Public Property Get CspKeyContainerInfo() As CspKeyContainerInfo
    Call VerifyLoaded
    
    Dim Params As CspParameters
    Set Params = Cor.NewCspParameters(mProviderType, mProviderName, mKeyContainerName)
    Params.KeyNumber = mKeyNumber
    Params.Flags = mFlags
    
    Set CspKeyContainerInfo = Cor.NewCspKeyContainerInfo(Params)
    CspKeyContainerInfo.RandomlyGenerated = mRandomlyGenerated
End Property

''
' Returns the name of the key exchange algorithm this instance represents.
'
' @Return The name of the key exchange algorithm.
' @remarks When an DSACryptoServiceProvider is created as a key exchange instead
' of a signature service, this returns 'DSA-PKCS1-KeyEx', otherwise nothing is returned.
'
Public Property Get KeyExchangeAlgorithm() As String
    ' return nothing
End Property

''
' Returns the DSA key size in bits.
'
' @return The key size in bits being used by this DSA instance.
' @remarks By default, the key size is 1024 for high encryption and 512 for base encryption providers.
'
Public Property Get KeySize() As Long
    KeySize = mKeySize
End Property

''
' Sets the key size.
'
' @param RHS The new key size.
' @remarks This property has no affect on the key size. To set the key size, use the
' Cor.NewDSACryptoServiceProvider constructor.
'
Public Property Let KeySize(ByVal RHS As Long)
    ' does nothing
End Property

''
' Returns an array of KeySizes objects defining the valid key sizes.
'
' @return An array of KeySizes objects.
' @remarks For high encryption systems, the legal key sizes go from 384 bits to 16384 bits
' in steps of 8 bits. For base encryption systems, the key sizes go from 384 bits to 512 bits
' in steps of 8 bits.
'
Public Property Get LegalKeySizes() As KeySizes()
    LegalKeySizes = mLegalKeySizes
End Property

''
' Returns if the current key will be stored in the current provider container.
'
' @return Returns True if the key will be stored, False otherwise.
' @remarks By default this is set to False. If a container is specified using the
' Cor.NewDSACryptoServiceProvider, then this will be set to True.<br>
' <br>A key is stored in the container to allow for each retrieval by any other
' service that knows where the key is. Only 1 key is stored in a container, therefore,
' if multiple providers share the same container, then the key may become invalid.
'
Public Property Get PersistKeyInCsp() As Boolean
    PersistKeyInCsp = mPersistKeyInCsp
End Property

''
' Sets if the current key will be stored in the current provider container.
'
' @param RHS Set to True to save the key on exit, False to have the key be deleted.
' @remarks By default this is set to False. If a container is specified using the
' Cor.NewDSACryptoServiceProvider, then this will be set to True.<br>
' <br>A key is stored in the container to allow for each retrieval by any other
' service that knows where the key is. Only 1 key is stored in a container, therefore,
' if multiple providers share the same container, then the key may become invalid.
'
Public Property Let PersistKeyInCsp(ByVal RHS As Boolean)
    mPersistKeyInCsp = RHS
End Property

''
' Returns if the key for this DSA instance is only the public half of the key pair.
'
' @return Returns True if only the public half of the key is present, False otherwise.
' @remarks DSA uses a private/public key pair to encrypt and decrypt. Only the public
' part is required to encrypt data. The private key is used to decrypt data.
'
Public Property Get PublicOnly() As Boolean
    Call VerifyLoaded
    PublicOnly = mPublicOnly
End Property

''
' Returns the signature algorithm.
'
' @return Always returns 'http://www.w3.org/2000/09/xmldsig#DSA-sha1'
'
Public Property Get SignatureAlgorithm() As String
    SignatureAlgorithm = "http://www.w3.org/2000/09/xmldsig#dsa-sha1"
End Property

''
' Releases any resources and disposes the DSA instance.
'
Public Sub Clear()
    Call CryptoAPI.DestroyKey(mKey)
    Call CryptoAPI.ReleaseContext(mProvider, Not mPersistKeyInCsp)
    
    mIsDisposed = True
End Sub

''
' Exports the key information as a CryptoAPI PRIVATEKEYBLOB or PUBLICKEYBLOB structure.
'
' @param IncludePrivateParameters A flag indicating if only the public key should be exported
' or is both the public and private keys should be exported.
' @return A exported key blob structure.
'
Public Function ExportCspBlob(ByVal IncludePrivateParameters As Boolean) As Byte()
    Call VerifyLoaded
    
    Dim BlobType As Long
    BlobType = IIf(IncludePrivateParameters, PRIVATEKEYBLOB, PUBLICKEYBLOB)
    
    Dim Size As Long
    If CryptExportKey(mKey, 0, BlobType, 0, ByVal 0&, Size) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
    
    Dim Blob() As Byte
    ReDim Blob(0 To Size - 1)
    
    If CryptExportKey(mKey, 0, BlobType, 0, Blob(0), Size) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
    
    ExportCspBlob = Blob
End Function

''
' Exports the DSA key parameters.
'
' @param IncludePrivateParameters A flag indicating if only the public key should be exported
' or is both the public and private keys should be exported.
' @return An DSAParameters object containing the parameters of the DSA algorithms.
'
Public Function ExportParameters(ByVal IncludePrivateParameters As Boolean) As DSAParameters
    Set ExportParameters = New DSAParameters
    
    If IncludePrivateParameters And (Not mPublicOnly) Then
        Call ExportParameters.FromCspBlob(Me.ExportCspBlob(False), Me.ExportCspBlob(True))
    Else
        Dim NullBytes() As Byte
        Call ExportParameters.FromCspBlob(Me.ExportCspBlob(False), NullBytes)
    End If
End Function

''
' Imports a DSA key in the form of an XML string.
'
' @param XmlString The XML string containing the key information.
'
Public Sub FromXmlString(ByVal XmlString As String)
    Dim Params As New DSAParameters
    Call Params.FromXmlString(XmlString)
    Call ImportParameters(Params)
End Sub

''
' Imports a CryptoAPI PRIVATEKEYBLOB or PUBLICKEYBLOB into the DSA provider.
'
' @param KeyBlob The blob key to be imported.
'
Public Sub ImportCspBlob(ByRef KeyBlob() As Byte)
    If cArray.IsNull(KeyBlob) Then _
        Throw Cor.NewArgumentNullException("KeyBlob", Environment.GetResourceString(ArgumentNull_Array))
    
    Call VerifyLoaded(False)
    Call DeleteKey
    
    mPublicOnly = (KeyBlob(LBound(KeyBlob)) = PUBLICKEYBLOB)
    
    If CryptImportKey(mProvider, KeyBlob(LBound(KeyBlob)), cArray.GetLength(KeyBlob), 0, IIf(mPublicOnly, 0, CRYPT_EXPORTABLE), mKey) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
End Sub

''
' Imports the DSA key parameters.
'
' @param Parameters The parameters to be imported as the new key.
'
Public Sub ImportParameters(ByVal Parameters As DSAParameters)
    If Parameters Is Nothing Then _
        Throw Cor.NewArgumentNullException("Parameters")
    
    Call ImportCspBlob(Parameters.ToCspBlob)
End Sub

''
' Creates a signature for a Stream object or Byte array.
'
' @param InputStreamOrBuffer The input source to be signed. This can be any Stream object or Byte array.
' @param Offset The starting index of the Byte array to begin processing. This is only applied to a Byte array.
' @param Count The number of bytes to be processed. This is only applied to a Byte array.
' @return A byte array representing the signature for the data.
'
Public Function SignData(ByRef InputStreamOrBuffer As Variant, Optional ByRef Offset As Variant, Optional ByRef Count As Variant) As Byte()
    Dim HashObj As New SHA1CryptoServiceProvider
    SignData = SignHash(HashObj.ComputeHash(InputStreamOrBuffer, Offset, Count), "sha1")
End Function

''
' Creates a signature for a hash value.
'
' @param RgbHash The hash to create a signature for.
' @param Str The type of hash used to create the hash value.
' @return A byte array representing the signature for the hash value.
' @remarks Only SHA1 hash values are supported.
'
Public Function SignHash(ByRef RgbHash() As Byte, ByVal Str As String) As Byte()
    Dim Hash As Long
    Hash = SetHash(RgbHash, Str)

    Dim Size As Long
    If CryptSignHash(Hash, AT_SIGNATURE, vbNullString, 0, ByVal 0&, Size) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))

    Dim Signature() As Byte
    ReDim Signature(0 To Size - 1)
    
    If CryptSignHash(Hash, AT_SIGNATURE, vbNullString, 0, Signature(0), Size) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
    
    Call CryptDestroyHash(Hash)
    
    ' Reverse the sections to match the .NET implementation
    Call cArray.Reverse(Signature, 0, 20)
    Call cArray.Reverse(Signature, 20, 20)
    
    SignHash = Signature
    Exit Function

errTrap:
    If Hash <> vbNullPtr Then Call CryptDestroyHash(Hash)
    Call Throw
End Function

''
' Exports the DSA key as an XML string.
'
' @param IncludePrivateParameters A flag indicating if the private portions of the key should be exported as well.
' @return An XML string representation of the DSA key.
'
Public Function ToXmlString(ByVal IncludePrivateParameters As Boolean) As String
    Dim Params As DSAParameters
    Set Params = Me.ExportParameters(IncludePrivateParameters)
    ToXmlString = Params.ToXmlString
End Function

''
' Verifies the signature for the specified data.
'
' @param RgbData The data to verify the signature of.
' @param Signature The signature to be verified.
' @return Returns True if the signature if valid for the data, False otherwise.
'
Public Function VerifyData(ByRef RgbData() As Byte, ByRef Signature() As Byte) As Boolean
    If cArray.IsNull(RgbData) Then _
        Throw Cor.NewArgumentNullException("Buffer", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim HashObj As New SHA1CryptoServiceProvider
    VerifyData = VerifyHash(HashObj.ComputeHash(RgbData), "SHA1", Signature)
End Function

''
' Verifies the signature for the specified hash value.
'
' @param RgbHash The hash value to verify the signature of.
' @param Str The type of hash used to create the hash value.
' @param Signature The signature to be verified.
' @return Returns True if the signature is valid for the hash value, False otherwise.
' @remarks Only SHA1 hash values are supported.
'
Public Function VerifyHash(ByRef RgbHash() As Byte, ByVal Str As String, ByRef Signature() As Byte) As Boolean
    If cArray.IsNull(Signature) Then _
        Throw Cor.NewArgumentNullException("Signature", Environment.GetResourceString(ArgumentNull_Array))
    
    Dim Hash As Long
    Hash = SetHash(RgbHash, Str)

    ' We need to reverse the two sections of the signature to deal with .NET's big-endian form.
    Dim RevSig() As Byte
    RevSig = Signature
    Call cArray.Reverse(RevSig, LBound(RevSig), 20)
    Call cArray.Reverse(RevSig, LBound(RevSig) + 20, 20)
    
    VerifyHash = CBool(CryptVerifySignature(Hash, RevSig(0), cArray.GetLength(RevSig), mKey, vbNullString, 0))
    Call CryptDestroyHash(Hash)
End Function

''
' Creates a new signature for the hash value.
'
' @param RgbHash The hash value to be signed.
' @return The hash value signature.
' @remarks Only SHA1 hash values are supported.
'
Public Function CreateSignature(ByRef RgbHash() As Byte) As Byte()
    CreateSignature = SignHash(RgbHash, "sha1")
End Function

''
' Verifies the signature for the hash value.
'
' @param RgbHash The hash value to verify the signature of.
' @param RgbSignature The signature to be verified.
' @return Returns True if the signature if valid, False otherwise.
' @remarks Only SHA1 hash values are supported.
'
Public Function VerifySignature(ByRef RgbHash() As Byte, ByRef RgbSignature() As Byte) As Boolean
    VerifySignature = VerifyHash(RgbHash, "sha1", RgbSignature)
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to test equality on.
' @return Boolean indicating equality.
' @see IObject
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
Friend Sub Init(ByVal KeySize As Long, ByVal Params As CspParameters)
    If KeySize > 0 Then mKeySize = KeySize
    
    If Not Params Is Nothing Then
        With Params
            mProviderName = .ProviderName
            mProviderType = .ProviderType
            mKeyContainerName = .KeyContainerName
            mKeyNumber = IIf(.KeyNumber = -1, DEF_KEYNUMBER, .KeyNumber)
            mFlags = .Flags
        End With
        
        mPersistKeyInCsp = True
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyLoaded(Optional ByVal IncludeKey As Boolean = True)
    If mIsDisposed Then _
        Throw New ObjectDisposedException
    
    Call LoadProvider
    
    If IncludeKey Then
        Call GenerateKey
    End If
End Sub

''
' This attempts to acquire a new provider and create a new key container
' within the specified provider. If the container name doesn't exist, then
' a random name will be created. If the container itself already exists,
' then the existing container will be used instead of creating a new one.
Private Sub LoadProvider()
    If mProvider <> vbNullPtr Then Exit Sub
    
    Dim Flags As Long
    If mFlags And CspProviderFlags.UseMachineKeyStore Then Flags = CRYPT_MACHINE_KEYSET
    
    mProvider = CryptoAPI.AcquireDSAContext(mKeyContainerName, mProviderName, mProviderType, Flags)
End Sub

Private Sub GenerateKey()
    If mKey <> vbNullPtr Then Exit Sub
    
    If mFlags And UseExistingKey Then
        ' Try to get an existing key.
        If CryptGetUserKey(mProvider, mKeyNumber, mKey) <> BOOL_FALSE Then Exit Sub
        
        ' Something went wrong. If it's something other than
        ' not finding the key, then it's bad, so fail.
        If Err.LastDllError <> NTE_NO_KEY Then _
            Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
            
        ' We didn't find an existing key, so continue and create a new one.
    End If
    
    Dim Flags As Long
    Flags = Helper.ShiftLeft(mKeySize, 16)  ' the key size is held in the upper 16-bits.
    
    ' We must specify to make the key archivable.
    If mFlags And UseArchivableKey Then Flags = Flags Or CRYPT_ARCHIVABLE
    
    ' We must specify to NOT create an exportable key.
    If (mFlags And UseNonExportableKey) = 0 Then Flags = Flags Or CRYPT_EXPORTABLE
    
    If Not CryptoAPI.SupportsAlgorithm(mProvider, CALG_DSS_SIGN) Then _
        Throw Cor.NewCryptographicException("The Cryptography Service Provider (CSP) does not support the specified algorithm.")

    ' Let'er rip! This can take a bit with large key sizes.
    If CryptGenKey(mProvider, mKeyNumber, Flags, mKey) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))
End Sub

Private Sub DeleteKey()
    If mKey <> vbNullPtr Then
        Call CryptDestroyKey(mKey)
        mKey = vbNullPtr
    End If
End Sub

Private Function SetHash(ByRef RgbHash() As Byte, ByVal Str As String) As Long
    Call VerifyLoaded
    
    If cArray.IsNull(RgbHash) Then _
        Throw Cor.NewArgumentNullException("RgbHash", Environment.GetResourceString(ArgumentNull_Array))
    
    Select Case LCase$(Str)
        Case "1.3.14.3.2.26", "sha", "sha1"
        Case Else
            Throw Cor.NewArgumentException("Hash OID not supported", "Str")
    End Select
    
    If cArray.GetLength(RgbHash) <> 20 Then _
        Throw Cor.NewCryptographicException("Invalid hash length.")
    
    On Error GoTo errTrap
    
    Dim Hash As Long
    If CryptCreateHash(mProvider, CALG_SHA1, 0, 0, Hash) = BOOL_FALSE Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))

    If CryptSetHashParam(Hash, HP_HASHVAL, RgbHash(LBound(RgbHash)), 0) = 0 Then _
        Throw Cor.NewCryptographicException(GetSystemMessage(Err.LastDllError))

    SetHash = Hash
    Exit Function
    
errTrap:
    If Hash <> vbNullPtr Then Call CryptDestroyHash(Hash)
    
    Throw
End Function

Private Function GetOID(ByVal HashObj As HashAlgorithm) As String
    Dim OID As String
    If TypeOf HashObj Is SHA1 Then
        OID = "1.3.14.3.2.26"
    End If
    
    GetOID = OID
End Function

Private Function GetHash(ByRef hAlg As Variant) As HashAlgorithm
    Dim HashObj As HashAlgorithm
    
    Select Case VarType(hAlg)
        Case vbString
            Select Case LCase$(hAlg)
                Case "1.3.14.3.2.26", "sha", "sha1":    Set HashObj = New SHA1CryptoServiceProvider
                Case Else
                    Throw Cor.NewArgumentException("Only SHA1 hash algorithm is supported.", "hAlg")
            End Select
            
        Case vbObject
            If hAlg Is Nothing Then _
                Throw Cor.NewArgumentNullException("hAlg")
            
            If TypeOf hAlg Is SHA1 Then
                Set HashObj = hAlg
            Else
                Throw Cor.NewArgumentException("Hash object must be SHA1 compatible.", "hAlg")
            End If
            
        Case Else
            Throw Cor.NewArgumentException("Invalid hash algorithm specified.", "hAlg")
    End Select
    
    Set GetHash = HashObj
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    ReDim mLegalKeySizes(0)
    Set mLegalKeySizes(0) = Cor.NewKeySizes(512, 1024, 64)
    
    mProviderType = DEF_PROVIDERTYPE
    mKeyNumber = DEF_KEYNUMBER
End Sub

Private Sub Class_Terminate()
    Call Clear
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   AsymmetricAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AsymmetricAlgorithm_Clear()
    Call Clear
End Sub

Private Function AsymmetricAlgorithm_Equals(Value As Variant) As Boolean
    AsymmetricAlgorithm_Equals = Equals(Value)
End Function

Private Sub AsymmetricAlgorithm_FromXmlString(ByVal XmlString As String)
    Call FromXmlString(XmlString)
End Sub

Private Function AsymmetricAlgorithm_GetHashCode() As Long
    AsymmetricAlgorithm_GetHashCode = GetHashCode
End Function

Private Property Get AsymmetricAlgorithm_KeyExchangeAlgorithm() As String
    AsymmetricAlgorithm_KeyExchangeAlgorithm = KeyExchangeAlgorithm
End Property

Private Property Let AsymmetricAlgorithm_KeySize(ByVal RHS As Long)
    KeySize = RHS
End Property

Private Property Get AsymmetricAlgorithm_KeySize() As Long
    AsymmetricAlgorithm_KeySize = KeySize
End Property

Private Property Get AsymmetricAlgorithm_LegalKeySizes() As KeySizes()
    AsymmetricAlgorithm_LegalKeySizes = LegalKeySizes
End Property

Private Property Get AsymmetricAlgorithm_SignatureAlgorithm() As String
    AsymmetricAlgorithm_SignatureAlgorithm = SignatureAlgorithm
End Property

Private Function AsymmetricAlgorithm_ToString() As String
    AsymmetricAlgorithm_ToString = ToString
End Function

Private Function AsymmetricAlgorithm_ToXmlString(ByVal IncludePrivateParameters As Boolean) As String
    AsymmetricAlgorithm_ToXmlString = ToXmlString(IncludePrivateParameters)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICspAsymmetricAlgorithm Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ICspAsymmetricAlgorithm_CspKeyContainerInfo() As CspKeyContainerInfo
    Set ICspAsymmetricAlgorithm_CspKeyContainerInfo = CspKeyContainerInfo
End Property

Private Function ICspAsymmetricAlgorithm_ExportCspBlob(ByVal IncludePrivateParameters As Boolean) As Byte()
    ICspAsymmetricAlgorithm_ExportCspBlob = ExportCspBlob(IncludePrivateParameters)
End Function

Private Sub ICspAsymmetricAlgorithm_ImportCspBlob(RawData() As Byte)
    Call ImportCspBlob(RawData)
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
'   DSA Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DSA_Clear()
    Call Clear
End Sub

Private Function DSA_Equals(Value As Variant) As Boolean
    DSA_Equals = Equals(Value)
End Function

Private Function DSA_ExportParameters(ByVal IncludePrivateParameters As Boolean) As DSAParameters
    Set DSA_ExportParameters = ExportParameters(IncludePrivateParameters)
End Function

Private Sub DSA_FromXmlString(ByVal XmlString As String)
    Call FromXmlString(XmlString)
End Sub

Private Function DSA_GetHashCode() As Long
    DSA_GetHashCode = GetHashCode
End Function

Private Sub DSA_ImportParameters(ByVal Parameters As DSAParameters)
    Call ImportParameters(Parameters)
End Sub

Private Property Get DSA_KeyExchangeAlgorithm() As String
    DSA_KeyExchangeAlgorithm = KeyExchangeAlgorithm
End Property

Private Property Let DSA_KeySize(ByVal RHS As Long)
    KeySize = RHS
End Property

Private Property Get DSA_KeySize() As Long
    DSA_KeySize = KeySize
End Property

Private Property Get DSA_LegalKeySizes() As KeySizes()
    DSA_LegalKeySizes = LegalKeySizes
End Property

Private Property Get DSA_SignatureAlgorithm() As String
    DSA_SignatureAlgorithm = SignatureAlgorithm
End Property

Private Function DSA_ToString() As String
    DSA_ToString = ToString
End Function

Private Function DSA_ToXmlString(ByVal IncludePrivateParameters As Boolean) As String
    DSA_ToXmlString = ToXmlString(IncludePrivateParameters)
End Function

Private Function DSA_CreateSignature(RgbHash() As Byte) As Byte()
    DSA_CreateSignature = CreateSignature(RgbHash)
End Function

Private Function DSA_VerifySignature(RgbHash() As Byte, RgbSignature() As Byte) As Boolean
    DSA_VerifySignature = VerifySignature(RgbHash, RgbSignature)
End Function



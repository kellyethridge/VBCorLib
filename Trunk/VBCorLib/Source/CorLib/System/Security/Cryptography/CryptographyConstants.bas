Attribute VB_Name = "CryptographyConstants"
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
' Module: CryptographyConstants
'
Option Explicit

Public Const PROV_RSA_FULL          As Long = 1
Public Const PROV_DSS_DH            As Long = 13

Public Const KP_IV                  As Long = 1
Public Const KP_MODE                As Long = 4
Public Const KP_MODE_BITS           As Long = 5
Public Const KP_EFFECTIVE_KEYLEN    As Long = 19
Public Const KP_SALT                As Long = 2
Public Const KP_PERMISSIONS         As Long = 6

Public Const PP_NAME                As Long = 4
Public Const PP_UNIQUE_CONTAINER    As Long = 36
Public Const PP_CONTAINER           As Long = 6
Public Const PP_PROVTYPE            As Long = 16
Public Const PP_ENUMALGS            As Long = 1


Public Const ALG_CLASS_DATA_ENCRYPT As Long = (3 * 2 ^ 13)
Public Const ALG_TYPE_BLOCK         As Long = (3 * 2 ^ 9)
Public Const ALG_SID_DES            As Long = 1
Public Const ALG_SID_RC2            As Long = 2
Public Const ALG_SID_3DES           As Long = 3
Public Const ALG_SID_3DES_112       As Long = 9
Public Const ALG_CLASS_HASH         As Long = (4 * 2 ^ 13)
Public Const ALG_TYPE_ANY           As Long = 0
Public Const ALG_SID_SHA1           As Long = 4
Public Const ALG_SID_MD5            As Long = 3
Public Const ALG_CLASS_KEY_EXCHANGE As Long = (5 * 2 ^ 13)
Public Const ALG_TYPE_RSA           As Long = (2 * 2 ^ 9)
Public Const ALG_SID_RSA_ANY        As Long = 0
Public Const ALG_CLASS_SIGNATURE    As Long = (1 * 2 ^ 13)
Public Const ALG_SID_DSS_ANY        As Long = 0
Public Const ALG_TYPE_DSS           As Long = (1 * 2 ^ 9)

Public Const CALG_DES               As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
Public Const CALG_RC2               As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2)
Public Const CALG_3DES              As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES)
Public Const CALG_3DES_112          As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES_112)
Public Const CALG_SHA1              As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1)
Public Const CALG_MD5               As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Public Const CALG_RSA_KEYX          As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Public Const CALG_RSA_SIGN          As Long = (ALG_CLASS_SIGNATURE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Public Const CALG_DSS_SIGN          As Long = (ALG_CLASS_SIGNATURE Or ALG_TYPE_DSS Or ALG_SID_DSS_ANY)

Public Const CRYPT_OAEP             As Long = &H40
Public Const CRYPT_EXPORTABLE       As Long = &H1
Public Const CRYPT_ARCHIVABLE       As Long = &H4000&
Public Const CRYPT_USER_PROTECTED   As Long = &H2
Public Const CRYPT_MODE_CBC         As Long = 1
Public Const CRYPT_MACHINE_KEYSET   As Long = &H20
Public Const CRYPT_NEWKEYSET        As Long = &H8
Public Const CRYPT_DELETEKEYSET     As Long = &H10
Public Const CRYPT_DECRYPT          As Long = &H2
Public Const CRYPT_OID_INFO_NAME_KEY As Long = 2
Public Const CRYPT_EXPORT           As Long = &H4
Public Const CRYPT_FIRST            As Long = 1
Public Const CRYPT_NO_SALT          As Long = &H10

Public Const PKCS7_PADDING          As Long = 2


Public Const HP_HASHSIZE            As Long = &H4
Public Const HP_HASHVAL             As Long = &H2

Public Const AT_KEYEXCHANGE         As Long = 1
Public Const AT_SIGNATURE           As Long = 2

Public Const NTE_BAD_KEYSET         As Long = &H80090016
Public Const NTE_EXISTS             As Long = &H8009000F
Public Const NTE_NO_KEY             As Long = &H8009000D

Public Const PUBLICKEYBLOB          As Long = &H6
Public Const PRIVATEKEYBLOB         As Long = &H7
Public Const SIMPLEBLOB             As Long = &H1

Public Const MS_ENH_DSS_DH_PROV     As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
Public Const MS_DEF_DSS_DH_PROV     As String = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
Public Const MS_DEF_PROV            As String = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_STRONG_PROV         As String = "Microsoft Strong Cryptographic Provider"
Public Const MS_ENHANCED_PROV       As String = "Microsoft Enhanced Cryptographic Provider v1.0"

Public Const SC_CLOSE               As Long = &HF060&
Public Const MF_BYCOMMAND           As Long = &H0&

Public Type PROV_ENUMALGS
    aiAlgid As Long
    dwBitLen As Long
    dwNameLen As Long
    szName As String * 20
End Type



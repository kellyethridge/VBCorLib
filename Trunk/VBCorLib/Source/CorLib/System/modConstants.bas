Attribute VB_Name = "modConstants"
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: modConstants
'
Option Explicit

Public Const PAGE_EXECUTE_READWRITE As Long = &H40

Public Const PICTYPE_ICON   As Long = 3
Public Const PICTYPE_BITMAP As Long = 1

Public Const vbMissing              As Long = vbError



Public Const MAX_PATH                   As Long = 260
Public Const MAX_DIRECTORY_PATH         As Long = 260
Public Const NO_ERROR                   As Long = 0


Public Const FILE_FLAG_OVERLAPPED       As Long = &H40000000
Public Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Public Const INVALID_HANDLE_VALUE       As Long = -1
Public Const FILE_TYPE_DISK             As Long = &H1
Public Const FILE_ATTRIBUTE_DIRECTORY   As Long = &H10
Public Const INVALID_FILE_ATTRIBUTES    As Long = -1

' File manipulation function attributes
Public Const GENERIC_READ               As Long = &H80000000
Public Const GENERIC_WRITE              As Long = &H40000000
Public Const OPEN_EXISTING              As Long = 3
Public Const PAGE_READONLY              As Long = &H2
Public Const PAGE_READWRITE             As Long = &H4
Public Const PAGE_WRITECOPY             As Long = &H8
Public Const INVALID_HANDLE             As Long = -1
Public Const FILE_SHARE_READ            As Long = 1
Public Const FILE_SHARE_WRITE           As Long = 2
Public Const STANDARD_RIGHTS_REQUIRED   As Long = &HF0000
Public Const SECTION_QUERY              As Long = &H1
Public Const SECTION_MAP_WRITE          As Long = &H2
Public Const SECTION_MAP_READ           As Long = &H4
Public Const SECTION_MAP_EXECUTE        As Long = &H8
Public Const SECTION_EXTEND_SIZE        As Long = &H10
Public Const SECTION_ALL_ACCESS         As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_READ              As Long = SECTION_MAP_READ
Public Const FILE_MAP_ALL_ACCESS        As Long = SECTION_ALL_ACCESS
Public Const NULL_HANDLE                As Long = 0

Public Const ERROR_PATH_NOT_FOUND       As Long = 3
Public Const ERROR_ACCESS_DENIED        As Long = 5
Public Const ERROR_FILE_NOT_FOUND       As Long = 2
Public Const ERROR_FILE_EXISTS          As Long = 80
Public Const ERROR_INSUFFICIENT_BUFFER  As Long = 122

' Registry Constants
Public Const REG_NONE                      As Long = 0
Public Const REG_UNKNOWN                   As Long = 0
Public Const REG_SZ                        As Long = 1
Public Const REG_DWORD                     As Long = 4
Public Const REG_BINARY                    As Long = 3
Public Const REG_MULTI_SZ                  As Long = 7
Public Const REG_EXPAND_SZ                 As Long = 2
Public Const REG_QWORD                     As Long = 11

Public Const ERROR_SUCCESS                 As Long = 0
Public Const ERROR_INVALID_HANDLE          As Long = 6
Public Const ERROR_INVALID_PARAMETER       As Long = 87
Public Const ERROR_CALL_NOT_IMPLEMENTED    As Long = 120
Public Const ERROR_MORE_DATA               As Long = 234
Public Const ERROR_NO_MORE_ITEMS           As Long = 259
Public Const ERROR_CANTOPEN                As Long = 1011
Public Const ERROR_CANTREAD                As Long = 1012
Public Const ERROR_CANTWRITE               As Long = 1013
Public Const ERROR_REGISTRY_RECOVERED      As Long = 1014
Public Const ERROR_REGISTRY_CORRUPT        As Long = 1015
Public Const ERROR_REGISTRY_IO_FAILED      As Long = 1016
Public Const ERROR_NOT_REGISTRY_FILE       As Long = 1017
Public Const ERROR_KEY_DELETED             As Long = 1018

' Locale Specifier
Public Const LOCALE_USER_DEFAULT = &H400

' GetCalendarInfo Constants
Public Const CAL_ITWODIGITYEARMAX   As Long = &H30
Public Const CAL_GREGORIAN          As Long = 1
Public Const CAL_HEBREW             As Long = 8
Public Const CAL_HIJRI              As Long = 6
Public Const CAL_JAPAN              As Long = 3
Public Const CAL_KOREA              As Long = 5
Public Const CAL_THAI               As Long = 7
Public Const CAL_TAIWAN             As Long = 4
Public Const CAL_RETURN_NUMBER      As Long = &H20000000


' Registry Root Keys
Public Const HKEY_CLASSES_ROOT      As Long = &H80000000
Public Const HKEY_CURRENT_CONFIG    As Long = &H80000005
Public Const HKEY_CURRENT_USER      As Long = &H80000001
Public Const HKEY_DYN_DATA          As Long = &H80000006
Public Const HKEY_LOCAL_MACHINE     As Long = &H80000002
Public Const HKEY_USERS             As Long = &H80000003
Public Const HKEY_PERFORMANCE_DATA  As Long = &H80000004

' Registry Flags
Public Const READ_CONTROL           As Long = &H20000
Public Const STANDARD_RIGHTS_ALL    As Long = &H1F0000
Public Const STANDARD_RIGHTS_READ   As Long = READ_CONTROL
Public Const KEY_QUERY_VALUE        As Long = &H1
Public Const KEY_SET_VALUE          As Long = &H2
Public Const KEY_CREATE_SUB_KEY     As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_CREATE_LINK        As Long = &H20
Public Const KEY_NOTIFY             As Long = &H10
Public Const SYNCHRONIZE            As Long = &H100000
Public Const KEY_READ               As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS         As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))



' Exception HResults
Public Const E_POINTER                  As Long = &H5B
Public Const COR_E_EXCEPTION            As Long = &H80131500
Public Const COR_E_SYSTEM               As Long = &H80131501
Public Const COR_E_RANK                 As Long = &H9
Public Const COR_E_INVALIDOPERATION     As Long = &H5
Public Const COR_E_INVALIDCAST          As Long = &HD
Public Const COR_E_INDEXOUTOFRANGE      As Long = &H9
Public Const COR_E_ARGUMENT             As Long = &H5
Public Const COR_E_ARGUMENTOUTOFRANGE   As Long = &H5
Public Const COR_E_OUTOFMEMORY          As Long = &H7
Public Const COR_E_FORMAT               As Long = &H80131537
Public Const COR_E_NOTSUPPORTED         As Long = &H1B6
Public Const COR_E_SERIALIZATION        As Long = &H14A
Public Const COR_E_ARRAYTYPEMISMATCH    As Long = &HD
Public Const COR_E_IO                   As Long = &H39
Public Const COR_E_FILENOTFOUND         As Long = &H35
Public Const COR_E_PLATFORMNOTSUPPORTED As Long = &H80131539
Public Const COR_E_PATHTOOLONG          As Long = &H800700CE
Public Const COR_E_DIRECTORYNOTFOUND    As Long = &H35
Public Const COR_E_ENDOFSTREAM          As Long = &H80070026
Public Const COR_E_ARITHMETIC           As Long = &H80070216
Public Const COR_E_OVERFLOW             As Long = &H6
Public Const COR_E_APPLICATION          As Long = &H80131600
Public Const COR_E_UNAUTHORIZEDACCESS   As Long = &H46
Public Const CORSEC_E_CRYPTO            As Long = &H80131430


'
' Cryptography Constants
'
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




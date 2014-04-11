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

Public Const vbMissing      As Long = vbError

' Ascii constants
Public Const vbTerminator       As Long = 0
Public Const vbUpperA           As Long = &H41
Public Const vbLowerA           As Long = &H61
Public Const vbLowerD           As Long = &H64
Public Const vbUpperD           As Long = &H44
Public Const vbLowerE           As Long = &H65
Public Const vbUpperE           As Long = &H45
Public Const vbLowerF           As Long = &H66
Public Const vbUpperF           As Long = &H46
Public Const vbLowerG           As Long = &H67
Public Const vbUpperG           As Long = &H47
Public Const vbLowerH           As Long = &H68
Public Const vbUpperH           As Long = &H48
Public Const vbLowerM           As Long = &H6D
Public Const vbUpperM           As Long = &H4D
Public Const vbLowerR           As Long = &H72
Public Const vbUpperR           As Long = &H52
Public Const vbLowerS           As Long = &H73
Public Const vbLowerT           As Long = &H74
Public Const vbUpperT           As Long = &H54
Public Const vbLowerU           As Long = &H75
Public Const vbUpperU           As Long = &H55
Public Const vbLowerY           As Long = &H79
Public Const vbUpperY           As Long = &H59
Public Const vbUpperZ           As Long = &HFA
Public Const vbLowerZ           As Long = &H7A
Public Const vbZero             As Long = &H30
Public Const vbOne              As Long = &H31
Public Const vbFive             As Long = &H35
Public Const vbEight            As Long = &H38
Public Const vbNine             As Long = &H39
Public Const vbPlus             As Long = &H2B
Public Const vbMinus            As Long = &H2D
Public Const vbBackSlash        As Long = &H5C
Public Const vbForwardSlash     As Long = &H2F
Public Const vbColon            As Long = &H3A
Public Const vbSemiColon        As Long = &H3B
Public Const vbEqual            As Long = &H3D
Public Const vbReturn           As Long = &HD
Public Const vbLineFeed         As Long = &HA
Public Const vbSpace            As Long = &H20
Public Const vbPound            As Long = &H23
Public Const vbDollar           As Long = &H24
Public Const vbPercent          As Long = &H25
Public Const vbDoubleQuote      As Long = &H22
Public Const vbSingleQuote      As Long = &H27
Public Const vbComma            As Long = &H2C
Public Const vbPeriod           As Long = &H2E
Public Const vbInvalidChar      As Long = &HFFFFFFFF

' Used for easy VarType comparison
Public Const vbIntegerArray     As Long = vbInteger Or vbArray
Public Const vbByteArray        As Long = vbByte Or vbArray
Public Const vbLongArray        As Long = vbLong Or vbArray
Public Const vbBooleanArray     As Long = vbBoolean Or vbArray
Public Const vbStringArray      As Long = vbString Or vbArray
Public Const vbVariantArray     As Long = vbVariant Or vbArray

' String versions
Public Const vbColonS           As String = ":"
Public Const vbSemiColonS       As String = ";"
Public Const vbBackSlashS       As String = "\"
Public Const vbForwardSlashS    As String = "/"
Public Const vbPeriodS          As String = "."


' SafeArray Constants
Public Const SIZEOF_SAFEARRAY               As Long = 16
Public Const SIZEOF_SAFEARRAYBOUND          As Long = 8
Public Const SIZEOF_SAFEARRAY1D             As Long = SIZEOF_SAFEARRAY + SIZEOF_SAFEARRAYBOUND
Public Const SIZEOF_GUID                    As Long = 16
Public Const SIZEOF_GUIDSAFEARRAY1D         As Long = SIZEOF_SAFEARRAY1D + SIZEOF_GUID

' Byte offsets into the SafeArray structure.
Public Const FFEATURES_OFFSET               As Long = 2
Public Const CBELEMENTS_OFFSET              As Long = 4
Public Const PVDATA_OFFSET                  As Long = 12
Public Const LBOUND_OFFSET                  As Long = 20
Public Const CLOCKS_OFFSET                  As Long = 8
Public Const CELEMENTS_OFFSET               As Long = 16

' Variant descriptions and offsets into the layout.
Public Const VARIANTDATA_OFFSET             As Long = 8
Public Const VT_BYREF                       As Long = &H4000
Public Const SIZEOF_VARIANT                 As Long = 16

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

' Constants used by CultureInfo and related classes when
' utilizing the CultureTable class.
Public Const LCID_INSTALLED                 As Long = &H1
Public Const LCID_SUPPORTED                 As Long = &H2
Public Const INVARIANT_LCID                 As Long = 127
             
Public Const ILCID                          As Long = 0
Public Const IPARENTLCID                    As Long = 1
Public Const ICALENDARTYPE                  As Long = 2
Public Const IFIRSTWEEKOFYEAR               As Long = 3
Public Const IFIRSTDAYOFWEEK                As Long = 4
Public Const ICURRENCYDECIMALDIGITS         As Long = 5
Public Const ICURRENCYNEGATIVEPATTERN       As Long = 6
Public Const ICURRENCYPOSITIVEPATTERN       As Long = 7
Public Const INUMBERDECIMALDIGITS           As Long = 8
Public Const INUMBERNEGATIVEPATTERN         As Long = 9
Public Const IPERCENTDECIMALDIGITS          As Long = 10
Public Const IPERCENTNEGATIVEPATTERN        As Long = 11
Public Const IPERCENTPOSITIVEPATTERN        As Long = 12


Public Const SENGLISHNAME                   As Long = 0
Public Const SDISPLAYNAME                   As Long = 1
Public Const SNAME                          As Long = 2
Public Const SNATIVENAME                    As Long = 3
Public Const STHREELETTERISOLANGUAGENAME    As Long = 4
Public Const STWOLETTERISOLANGUAGENAME      As Long = 5
Public Const STHREELETTERWINDOWSLANGUAGENAME As Long = 6
Public Const SOPTIONALCALENDARS             As Long = 7
Public Const SABBREVIATEDDAYNAMES           As Long = 8
Public Const SABBREVIATEDMONTHNAMES         As Long = 9
Public Const SAMDESIGNATOR                  As Long = 10
Public Const SDATESEPARATOR                 As Long = 11
Public Const SDAYNAMES                      As Long = 12
Public Const SLONGDATEPATTERN               As Long = 13
Public Const SLONGTIMEPATTERN               As Long = 14
Public Const SMONTHDAYPATTERN               As Long = 15
Public Const SMONTHNAMES                    As Long = 16
Public Const SPMDESIGNATOR                  As Long = 17
Public Const SSHORTDATEPATTERN              As Long = 18
Public Const SSHORTTIMEPATTERN              As Long = 19
Public Const STIMESEPARATOR                 As Long = 20
Public Const SYEARMONTHPATTERN              As Long = 21
Public Const SALLLONGDATEPATTERNS           As Long = 22
Public Const SALLSHORTDATEPATTERNS          As Long = 23
Public Const SALLLONGTIMEPATTERNS           As Long = 24
Public Const SALLSHORTTIMEPATTERNS          As Long = 25
Public Const SALLMONTHDAYPATTERNS           As Long = 26
Public Const SCURRENCYGROUPSIZES            As Long = 27
Public Const SNUMBERGROUPSIZES              As Long = 28
Public Const SPERCENTGROUPSIZES             As Long = 29
Public Const SCURRENCYDECIMALSEPARATOR      As Long = 30
Public Const SCURRENCYGROUPSEPARATOR        As Long = 31
Public Const SCURRENCYSYMBOL                As Long = 32
Public Const SNANSYMBOL                     As Long = 33
Public Const SNEGATIVEINFINITYSYMBOL        As Long = 34
Public Const SNEGATIVESIGN                  As Long = 35
Public Const SNUMBERDECIMALSEPARATOR        As Long = 36
Public Const SNUMBERGROUPSEPARATOR          As Long = 37
Public Const SPERCENTDECIMALSEPARATOR       As Long = 38
Public Const SPERCENTGROUPSEPARATOR         As Long = 39
Public Const SPERCENTSYMBOL                 As Long = 40
Public Const SPERMILLESYMBOL                As Long = 41
Public Const SPOSITIVEINFINITYSYMBOL        As Long = 42
Public Const SPOSITIVESIGN                  As Long = 43


' Used for GetLocaleInfo API
Public Const LOCALE_RETURN_NUMBER           As Long = &H20000000
Public Const LOCALE_ICENTURY                As Long = &H24
Public Const LOCALE_ICOUNTRY                As Long = &H5
Public Const LOCALE_ICURRDIGITS             As Long = &H19
Public Const LOCALE_ICURRENCY               As Long = &H1B
Public Const LOCALE_IDATE                   As Long = &H21
Public Const LOCALE_IDAYLZERO               As Long = &H26
Public Const LOCALE_IDEFAULTANSICODEPAGE    As Long = &H1004
Public Const LOCALE_IDEFAULTCODEPAGE        As Long = &HB
Public Const LOCALE_IDEFAULTCOUNTRY         As Long = &HA
Public Const LOCALE_IDEFAULTEBCDICCODEPAGE  As Long = &H1012
Public Const LOCALE_IDEFAULTLANGUAGE        As Long = &H9
Public Const LOCALE_IDEFAULTMACCODEPAGE     As Long = &H1011
Public Const LOCALE_IDIGITS                 As Long = &H11
Public Const LOCALE_IDIGITSUBSTITUTION      As Long = &H1014
Public Const LOCALE_IFIRSTDAYOFWEEK         As Long = &H100C
Public Const LOCALE_IFIRSTWEEKOFYEAR        As Long = &H100D
Public Const LOCALE_IINTLCURRDIGITS         As Long = &H1A
Public Const LOCALE_ILANGUAGE               As Long = &H1
Public Const LOCALE_ILDATE                  As Long = &H22
Public Const LOCALE_ILZERO                  As Long = &H12
Public Const LOCALE_IMEASURE                As Long = &HD
Public Const LOCALE_IMONLZERO               As Long = &H27
Public Const LOCALE_INEGCURR                As Long = &H1C
Public Const LOCALE_INEGNUMBER              As Long = &H1010
Public Const LOCALE_INEGSEPBYSPACE          As Long = &H57
Public Const LOCALE_INEGSIGNPOSN            As Long = &H53
Public Const LOCALE_INEGSYMPRECEDES         As Long = &H56
Public Const LOCALE_IOPTIONALCALENDAR       As Long = &H100B
Public Const LOCALE_IPAPERSIZE              As Long = &H100A
Public Const LOCALE_IPOSSEPBYSPACE          As Long = &H55
Public Const LOCALE_IPOSSIGNPOSN            As Long = &H52
Public Const LOCALE_IPOSSYMPRECEDES         As Long = &H54
Public Const LOCALE_ITIME                   As Long = &H23
Public Const LOCALE_ITIMEMARKPOSN           As Long = &H1005
Public Const LOCALE_ITLZERO                 As Long = &H25
Public Const LOCALE_NOUSEROVERRIDE          As Long = &H80000000
Public Const LOCALE_S1159                   As Long = &H28
Public Const LOCALE_S2359                   As Long = &H29
Public Const LOCALE_SABBREVCTRYNAME         As Long = &H7
Public Const LOCALE_SABBREVDAYNAME1         As Long = &H31
Public Const LOCALE_SABBREVDAYNAME2         As Long = &H32
Public Const LOCALE_SABBREVDAYNAME3         As Long = &H33
Public Const LOCALE_SABBREVDAYNAME4         As Long = &H34
Public Const LOCALE_SABBREVDAYNAME5         As Long = &H35
Public Const LOCALE_SABBREVDAYNAME6         As Long = &H36
Public Const LOCALE_SABBREVDAYNAME7         As Long = &H37
Public Const LOCALE_SABBREVLANGNAME         As Long = &H3
Public Const LOCALE_SABBREVMONTHNAME1       As Long = &H44
Public Const LOCALE_SABBREVMONTHNAME10      As Long = &H4D
Public Const LOCALE_SABBREVMONTHNAME11      As Long = &H4E
Public Const LOCALE_SABBREVMONTHNAME12      As Long = &H4F
Public Const LOCALE_SABBREVMONTHNAME13      As Long = &H100F
Public Const LOCALE_SABBREVMONTHNAME2       As Long = &H45
Public Const LOCALE_SABBREVMONTHNAME3       As Long = &H46
Public Const LOCALE_SABBREVMONTHNAME4       As Long = &H47
Public Const LOCALE_SABBREVMONTHNAME5       As Long = &H48
Public Const LOCALE_SABBREVMONTHNAME6       As Long = &H49
Public Const LOCALE_SABBREVMONTHNAME7       As Long = &H4A
Public Const LOCALE_SABBREVMONTHNAME8       As Long = &H4B
Public Const LOCALE_SABBREVMONTHNAME9       As Long = &H4C
Public Const LOCALE_SCOUNTRY                As Long = &H6
Public Const LOCALE_SCURRENCY               As Long = &H14
Public Const LOCALE_SDATE                   As Long = &H1D
Public Const LOCALE_SDAYNAME1               As Long = &H2A
Public Const LOCALE_SDAYNAME2               As Long = &H2B
Public Const LOCALE_SDAYNAME3               As Long = &H2C
Public Const LOCALE_SDAYNAME4               As Long = &H2D
Public Const LOCALE_SDAYNAME5               As Long = &H2E
Public Const LOCALE_SDAYNAME6               As Long = &H2F
Public Const LOCALE_SDAYNAME7               As Long = &H30
Public Const LOCALE_SDECIMAL                As Long = &HE
Public Const LOCALE_SENGCOUNTRY             As Long = &H1002
Public Const LOCALE_SENGCURRNAME            As Long = &H1007
Public Const LOCALE_SENGLANGUAGE            As Long = &H1001
Public Const LOCALE_SGROUPING               As Long = &H10
Public Const LOCALE_SINTLSYMBOL             As Long = &H15
Public Const LOCALE_SISO3166CTRYNAME        As Long = &H5A
Public Const LOCALE_SISO639LANGNAME         As Long = &H59
Public Const LOCALE_SLANGUAGE               As Long = &H2
Public Const LOCALE_SLIST                   As Long = &HC
Public Const LOCALE_SLONGDATE               As Long = &H20
Public Const LOCALE_SMONDECIMALSEP          As Long = &H16
Public Const LOCALE_SMONGROUPING            As Long = &H18
Public Const LOCALE_SMONTHNAME1             As Long = &H38
Public Const LOCALE_SMONTHNAME10            As Long = &H41
Public Const LOCALE_SMONTHNAME11            As Long = &H42
Public Const LOCALE_SMONTHNAME12            As Long = &H43
Public Const LOCALE_SMONTHNAME13            As Long = &H100E
Public Const LOCALE_SMONTHNAME2             As Long = &H39
Public Const LOCALE_SMONTHNAME3             As Long = &H3A
Public Const LOCALE_SMONTHNAME4             As Long = &H3B
Public Const LOCALE_SMONTHNAME5             As Long = &H3C
Public Const LOCALE_SMONTHNAME6             As Long = &H3D
Public Const LOCALE_SMONTHNAME7             As Long = &H3E
Public Const LOCALE_SMONTHNAME8             As Long = &H3F
Public Const LOCALE_SMONTHNAME9             As Long = &H40
Public Const LOCALE_SMONTHOUSANDSEP         As Long = &H17
Public Const LOCALE_SNATIVECTRYNAME         As Long = &H8
Public Const LOCALE_SNATIVECURRNAME         As Long = &H1008
Public Const LOCALE_SNATIVEDIGITS           As Long = &H13
Public Const LOCALE_SNATIVELANGNAME         As Long = &H4
Public Const LOCALE_SNEGATIVESIGN           As Long = &H51
Public Const LOCALE_SPOSITIVESIGN           As Long = &H50
Public Const LOCALE_SSHORTDATE              As Long = &H1F
Public Const LOCALE_SSORTNAME               As Long = &H1013
Public Const LOCALE_STHOUSAND               As Long = &HF&
Public Const LOCALE_STIME                   As Long = &H1E
Public Const LOCALE_STIMEFORMAT             As Long = &H1003
Public Const LOCALE_SYEARMONTH              As Long = &H1006

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




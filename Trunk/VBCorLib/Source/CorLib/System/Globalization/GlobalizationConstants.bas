Attribute VB_Name = "GlobalizationConstants"
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
' Module: GlobalizationConstants
'
Option Explicit

Public Const NumberFormatInfoTypeName   As String = "NumberFormatInfo"
Public Const DateTimeFormatInfoTypeName As String = "DateTimeFormatInfo"


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


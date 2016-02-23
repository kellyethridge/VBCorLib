Attribute VB_Name = "Statics"
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
' Module: Statics
'

''
' This is the set of classes that provide methods without the need to
' explicitly instanciate an object. They are shared by the entire application
' and exposed to client applications through the StaticClasses class.
'
Option Explicit

Public Type NullVersionableCollection
    Instance As New IVersionableCollection
End Type

Public Cor                      As New Constructors
Public Object                   As New ObjectStatic
Public CorArray                 As New CorArray
Public CorString                As New CorString
Public Comparer                 As New ComparerStatic
Public CaseInsensitiveComparer  As New CaseInsensitiveComparerStatic
Public Environment              As New Environment
Public Buffer                   As New Buffer
Public NumberFormatInfo         As New NumberFormatInfoStatic
Public BitConverter             As New BitConverter
Public TimeSpan                 As New TimeSpanStatic
Public CorDateTime              As New CorDateTimeStatic
Public DateTimeFormatInfo       As New DateTimeFormatInfoStatic
Public CultureTable             As New CultureTable
Public CultureInfo              As New CultureInfoStatic
Public Path                     As New Path
Public Encoding                 As New EncodingStatic
Public Directory                As New Directory
Public File                     As New File
Public Console                  As New Console
Public Calendar                 As New CalendarStatic
Public GregorianCalendar        As New GregorianCalendarStatic
Public JulianCalendar           As New JulianCalendarStatic
Public HebrewCalendar           As New HebrewCalendarStatic
Public KoreanCalendar           As New KoreanCalendarStatic
Public ThaiBuddhistCalendar     As New ThaiBuddhistCalendarStatic
Public HijriCalendar            As New HijriCalendarStatic
Public ArrayList                As New ArrayListStatic
Public Version                  As New VersionStatic
Public BitArray                 As New BitArrayStatic
Public TimeZone                 As New TimeZoneStatic
Public Stream                   As New StreamStatic
Public TextReader               As New TextReaderStatic
Public Registry                 As New Registry
Public RegistryKey              As New RegistryKeyStatic
Public Guid                     As New GuidStatic
Public Convert                  As New Convert
Public ResourceManager          As New ResourceManagerStatic
Public DriveInfo                As New DriveInfoStatic
Public CorMath                  As New CorMath
Public EventArgs                As New EventArgsStatic
Public DES                      As New DESStatic
Public TripleDES                As New TripleDESStatic
Public RC2                      As New RC2Static
Public Rijndael                 As New RijndaelStatic
Public CryptoConfig             As New CryptoConfig
Public StopWatch                As New StopWatchStatic
Public MD5                      As New MD5Static
Public SHA1                     As New SHA1Static
Public SHA256                   As New SHA256Static
Public SHA512                   As New SHA512Static
Public SHA384                   As New SHA384Static
Public MACTripleDES             As New MACTripleDESStatic
Public HMAC                     As New HMACStatic
Public RSA                      As New RSAStatic
Public SecurityElement          As New SecurityElementStatic
Public CryptoAPI                As New CryptoAPI
Public CryptoHelper             As New CryptoHelper
Public BigInteger               As New BigIntegerStatic
Public Thread                   As New ThreadStatic
Public StringBuilderCache       As New StringBuilderCache
Public EqualityComparer         As New EqualityComparerStatic
Public MyBase                   As New ObjectBase
Public Char                     As New Char
Public Error                    As New ErrorStatic
Public IOError                  As New IOError
Public Number                   As New NumberStatic
Public NullVersionableCollection As NullVersionableCollection
Public StringComparer           As New StringComparerStatic
Public RSACryptoServiceProvider As New RSACryptoServiceProviderStatic



Attribute VB_Name = "modStaticClasses"
'    CopyRight (c) 2005 Kelly Ethridge
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
'    Module: modStaticClasses
'

''
' This is the set of classes that provide methods without the need to
' explicitly instanciate an object. They are shared by the entire application
' and exposed to client applications through the StaticClasses class.
'
Option Explicit

' these are defined As New so they will only be created when first used.
Public Cor                      As New Constructors
Public Object                   As New ObjectStatic
Public cArray                   As New cArray
Public cString                  As New cString
Public Comparer                 As New ComparerStatic
Public Environment              As New Environment
Public Buffer                   As New Buffer
Public NumberFormatInfo         As New NumberFormatInfoStatic
Public BitConverter             As New BitConverter
Public TimeSpan                 As New TimeSpanStatic
Public cDateTime                As New cDateTimeStatic
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
Public MathExt                  As New MathExt
Public EventArgs                As New EventArgsStatic
Public DES                      As New DESStatic
Public TripleDES                As New TripleDESStatic
Public RC2                      As New RC2Static
Public Rijndael                 As New RijndaelStatic
Public CryptoConfig             As New CryptoConfig
Public SHA1                     As New SHA1Static
Public StopWatch                As New StopWatchStatic
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
Public StringBuilders           As New StringBuilderCache
Public EqualityComparer         As New EqualityComparerStatic
Public MyBase                   As New ObjectBase
Public Char                     As New Char



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticClasses"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'    Module: StaticClasses
'

''
' Provides direct access to static class methods.
'
' @remarks VB6 does not support static class methods. To replicate the syntax of
' accessing static methods, a function of the same name as the class that is
' supposed to contain the static methods is defined in this class. This class
' has instancing of <i>GlobalMultiUse</i>. This means that just the name of these
' functions can be typed directly without having to instantiate the <b>PublicFunctions</b>
' class. By just typing in any of these function names, the underlying class which
' contains the static methods is returned, and those methods are then available
' for calling. This gives the illusion of static methods being called on a class.
' <p>When typing in one of these function names, it should be followed immediately
' by a period to allow access to the static methods of that class.</p>
' <pre>
' '' Print the name of the current user
' Debug.Print Environment.UserName
' </pre>
' This small example simply refers to the <i>Environment</i> function in this
' class, then directly calls a method on the returned object from the
' <i>Environment</i> function.
'
Option Explicit

''
' Represents the sign values of BigInteger values.
'
' @param Negative The value is a negative number.
' @param Zero The value is zero.
' @param Positive The value is a positive number.
'
Public Enum Sign
    Negative = -1
    Zero = 0
    Positive = 1
End Enum


''
' Returns a object containing all of the parameterized constructors
' available within the VBCorLib library.
'
' @return A object containing all of the constructors in VBCorLib.
' @see constructors
'
Public Function Cor() As Constructors
    Set Cor = modStaticClasses.Cor
End Function

''
' Returns an object used to access the <b>cArray</b> methods
' without instantiating the <b>cArray</b> class.
'
' @return A reference to the <b>cArray</b> object used to
' perform functions on arrays.
' @see cArray
'
Public Function cArray() As cArray
    Set cArray = modStaticClasses.cArray
End Function

''
' Returns an object used to access the <b>cString</b> methods
' without instantiating the <b>cString</b> class.
'
' @return A reference to the <b>cString</b> object used to
' perform functions on Strings.
' @see cString
'
Public Function cString() As cString
    Set cString = modStaticClasses.cString
End Function

''
' Returns an object used to access the <b>Comparer</b> static functions.
'
' @return A reference to the object used to access the <b>Comparer</b> static functions.
' @see Comparer
' @see ComparerStatic
'
Public Function Comparer() As ComparerStatic
    Set Comparer = modStaticClasses.Comparer
End Function

''
' Returns a reference used to access the methods of the <b>Environment</b> functions.
'
' @return A reference that can be used to access methods of the <b>Environment</b> class.
' @see Environment
'
Public Function Environment() As Environment
    Set Environment = modStaticClasses.Environment
End Function

''
' Returns a reference used to access static methods of the <b>BitArray</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitArray</b> class.
' @see BitArray
' @see BitArrayStatic
Public Function BitArray() As BitArrayStatic
    Set BitArray = modStaticClasses.BitArray
End Function

''
' Returns a reference used to access static methods of the <b>Buffer</b> class.
'
' @return A reference that can be used to access static methods of the <b>Buffer</b> class.
' @see Buffer
'
Public Function Buffer() As Buffer
    Set Buffer = modStaticClasses.Buffer
End Function

''
' Returns a reference used to access static methods of the <b>NumberFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>NumberFormatInfo</b> class.
' @see NumberFormatInfo
' @see NumberFormatInfoStatic
'
Public Function NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = modStaticClasses.NumberFormatInfo
End Function

''
' Returns a reference used to access static methods of the <b>BitConverter</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitConverter</b> class.
' @see BitConverter
'
Public Function BitConverter() As BitConverter
    Set BitConverter = modStaticClasses.BitConverter
End Function

''
' Returns a reference used to access static methods of the <b>Version</b> class.
'
' @return A reference that can be used to access static methods of the <b>Version</b> class.
' @see Version
' @see VersionStatic
'
Public Function Version() As VersionStatic
    Set Version = modStaticClasses.Version
End Function

''
' Returns a reference used to access static methods of the <b>TimeSpan</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeSpan</b> class.
' @see TimeSpan
' @see TimeSpanStatic
'
Public Function TimeSpan() As TimeSpanStatic
    Set TimeSpan = modStaticClasses.TimeSpan
End Function

''
' Returns a reference used to access static methods of the <b>cDateTime</b> class.
'
' @return A reference that can be used to access static methods of the <b>cDateTime</b> class.
' @see cDateTime
' @see cDateTimeStatic
'
Public Function cDateTime() As cDateTimeStatic
    Set cDateTime = modStaticClasses.cDateTime
End Function

''
' Returns a reference used to access static methods of the <b>DateTimeFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>DateTimeFormatInfo</b> class.
' @see DateTimeFormatInfo
' @see DateTimeFormatInfoStatic
'
Public Function DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = modStaticClasses.DateTimeFormatInfo
End Function

''
' Returns a reference used to access static methods of the <b>CultureInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>CultureInfo</b> class.
' @see CultureInfo
' @see CultureInfoStatic
'
Public Function CultureInfo() As CultureInfoStatic
    Set CultureInfo = modStaticClasses.CultureInfo
End Function

''
' Returns a reference used to access static methods of the <b>TimeZone</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeZone</b> class.
' @see TimeZone
' @see TimeZoneStatic
'
Public Function TimeZone() As TimeZoneStatic
    Set TimeZone = modStaticClasses.TimeZone
End Function

''
' Returns a reference used to access static methods of the <b>Path</b> class.
'
' @return A reference that can be used to access static methods of the <b>Path</b> class.
' @see Path
'
Public Function Path() As Path
    Set Path = modStaticClasses.Path
End Function

''
' Returns a reference used to access static methods of the <b>Encoding</b> class.
'
' @return A reference that can be used to access static methods of the <b>Encoding</b> class.
' @see Encoding
' @see EncodingStatic
'
Public Function Encoding() As EncodingStatic
    Set Encoding = modStaticClasses.Encoding
End Function

''
' Returns a reference used to access static methods of the <b>TextReader</b> class.
'
' @return A reference that can be used to access static methods of the <b>TextReader</b> class.
' @see TextReader
' @see TextReaderStatic
' @see StreamReader
'
Public Function TextReader() As TextReaderStatic
    Set TextReader = modStaticClasses.TextReader
End Function

''
' Returns a reference used to access static methods of the <b>Directory</b> class.
'
' @return A reference that can be used to access static methods of the <b>Directory</b> class.
' @see Directory
' @see DirectoryInfo
' @see File
' @see FileInfo
' @see DriveInfo
' @see DriveInfoStatic
'
Public Function Directory() As Directory
    Set Directory = modStaticClasses.Directory
End Function

''
' Returns a reference used to access static methods of the <b>File</b> class.
'
' @return A reference that can be used to access static methods of the <b>File</b> class.
' @see File
' @see FileInfo
' @see Directory
' @see DirectoryInfo
' @see DriveInfo
' @see DriveInfoStatic
'
Public Function File() As File
    Set File = modStaticClasses.File
End Function

''
' Returns a reference used to access static methods of the <b>Stream</b> class.
'
' @return A reference that can be used to access static methods of the <b>Stream</b> class.
' @see Stream
' @see StreamStatic
'
Public Function Stream() As StreamStatic
    Set Stream = modStaticClasses.Stream
End Function

''
' Returns a reference used to access static methods of the <b>GregorianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>GregorianCalendar</b> class.
' @see GregorianCalendar
' @see GregorianCalendarStatic
'
Public Function GregorianCalendar() As GregorianCalendarStatic
    Set GregorianCalendar = modStaticClasses.GregorianCalendar
End Function

''
' Returns a reference used to access static methods of the <b>JulianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>JulianCalendar</b> class.
' @see JulianCalendar
' @see JulianCalendarStatic
'
Public Function JulianCalendar() As JulianCalendarStatic
    Set JulianCalendar = modStaticClasses.JulianCalendar
End Function

''
' Returns a reference used to access static methods of the <b>HebrewCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HebrewCalendar</b> class.
' @see HebrewCalendar
' @see HebrewCalendarStatic
'
Public Function HebrewCalendar() As HebrewCalendarStatic
    Set HebrewCalendar = modStaticClasses.HebrewCalendar
End Function

''
' Returns a reference used to access static methods of the <b>KoreanCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>KoreanCalendar</b> class.
' @see KoreanCalendar
' @see KoreanCalendarStatic
'
Public Function KoreanCalendar() As KoreanCalendarStatic
    Set KoreanCalendar = modStaticClasses.KoreanCalendar
End Function

''
' Returns a reference used to access static methods of the <b>HijriCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HijriCalendar</b> class.
' @see HijriCalendar
' @see HijriCalendarStatic
'
Public Function HijriCalendar() As HijriCalendarStatic
    Set HijriCalendar = modStaticClasses.HijriCalendar
End Function

''
' Returns a reference used to access static methods of the <b>Calendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>Calendar</b> class.
' @see Calendar
' @see CalendarStatic
'
Public Function Calendar() As CalendarStatic
    Set Calendar = modStaticClasses.Calendar
End Function

''
' Returns a reference used to access static methods of the <b>Console</b> class.
'
' @return A reference that can be used to access static methods of the <b>Console</b> class.
' @see Console
'
Public Function Console() As Console
    Set Console = modStaticClasses.Console
End Function

''
' Returns a reference used to access static methods of the <b>ArrayList</b> class.
'
' @return A reference that can be used to access static methods of the <b>ArrayList</b> class.
' @see ArrayList
' @see ArrayListStatic
'
Public Function ArrayList() As ArrayListStatic
    Set ArrayList = modStaticClasses.ArrayList
End Function

''
' Returns a reference used to access static methods of the <b>Registry</b> class.
'
' @return A reference that can be used to access static methods of the <b>Registry</b> class.
' @see Registry
' @see RegistryKey
'
Public Function Registry() As Registry
    Set Registry = modStaticClasses.Registry
End Function

''
' Returns a reference used to access static methods of the <b>RegistryKey</b> class.
'
' @return A reference that can be used to access static methods of the <b>RegistryKey</b> class.
' @see RegistryKey
' @see RegistryKeyStatic
' @see Registry
'
Public Function RegistryKey() As RegistryKeyStatic
    Set RegistryKey = modStaticClasses.RegistryKey
End Function

''
' Provides static methods for creating new <b>Guid</b> objects.
'
' @return A reference that can be used to access the static methods of the <b>Guid</b> class.
' @see GuidStatic
' @see Guid
'
Public Function Guid() As GuidStatic
    Set Guid = modStaticClasses.Guid
End Function

''
' Provides methods for converting values.
'
' @return A reference to the method provider.
' @see Convert
'
Public Function Convert() As Convert
    Set Convert = modStaticClasses.Convert
End Function

''
' Provides methods for creating custom resoure managers.
'
' @return A reference to the method provider.
' @see ResourceManagerStatic
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
'
Public Function ResourceManager() As ResourceManagerStatic
    Set ResourceManager = modStaticClasses.ResourceManager
End Function

''
' Provides methods for retrieve information about all drives on the local machine.
'
' @return A reference to the method provider
' @see DriveInfo
' @see DirectoryInfo
' @see Directory
' @see FileInfo
' @see File
'
Public Function DriveInfo() As DriveInfoStatic
    Set DriveInfo = modStaticClasses.DriveInfo
End Function

''
' Provides default methods for classes that implement the IObject interface.
'
' @return A reference to default implementations of the IObject interface.
'
Public Function Object() As ObjectStatic
    Set Object = modStaticClasses.Object
End Function

''
' Returns a reference used to access the DES static methods.
'
' @return A reference to the single DESStatic instance.
'
Public Function DES() As DESStatic
    Set DES = modStaticClasses.DES
End Function

''
' Returns a reference used to access the TripleDES static methods.
'
' @return A reference to the single TripleDESStatic instance.
'
Public Function TripleDES() As TripleDESStatic
    Set TripleDES = modStaticClasses.TripleDES
End Function

''
' Returns a reference used to access the RC2 static methods.
'
' @return A reference to the single RC2Static instance.
'
Public Function RC2() As RC2Static
    Set RC2 = modStaticClasses.RC2
End Function

''
' Returns a reference used to access the Rijndael static methods.
'
' @return A reference to the single RijndaelStatic instance.
'
Public Function Rijndael() As RijndaelStatic
    Set Rijndael = modStaticClasses.Rijndael
End Function

''
' Returns a reference used to access the CryptoConfig static methods.
'
' @return A reference to the single CryptoConfig instance.
'
Public Function CryptoConfig() As CryptoConfig
    Set CryptoConfig = modStaticClasses.CryptoConfig
End Function

''
' Returns a reference used to access the SHA1 static methods.
'
' @return A reference to the single SHA1Static instance.
'
Public Function SHA1() As SHA1Static
    Set SHA1 = modStaticClasses.SHA1
End Function

''
' Returns a reference used to access the StopWatch static methods.
'
' @return A reference to the single StopWatchStatic instance.
'
Public Function StopWatch() As StopWatchStatic
    Set StopWatch = modStaticClasses.StopWatch
End Function

''
' Returns a reference used to access the SHA256 static methods.
'
' @return A reference to the single SHA256Static instance.
'
Public Function SHA256() As SHA256Static
    Set SHA256 = modStaticClasses.SHA256
End Function

''
' Returns a reference used to access the SHA512 static methods.
'
' @return A reference to the single SHA512Static instance.
'
Public Function SHA512() As SHA512Static
    Set SHA512 = modStaticClasses.SHA512
End Function

''
' Returns a reference used to access the SHA384 static methods.
'
' @return A reference to the single SHA384Static instance.
'
Public Function SHA384() As SHA384Static
    Set SHA384 = modStaticClasses.SHA384
End Function

''
' Returns a reference used to access the MACTripleDES static methods.
'
' @return A reference to the single MACTripleDESStatic instance.
'
Public Function MACTripleDES() As MACTripleDESStatic
    Set MACTripleDES = modStaticClasses.MACTripleDES
End Function

''
' Returns a reference used to access the HMAC static methods.
'
' @return A reference to the single HMACStatic instance.
'
Public Function HMAC() As HMACStatic
    Set HMAC = modStaticClasses.HMAC
End Function

''
' Returns a reference used to access the RSA static methods.
'
' @return A reference to the single RSAStatic instance.
'
Public Function RSA() As RSAStatic
    Set RSA = modStaticClasses.RSA
End Function

''
' Returns a reference used to access the SecurityElement static methods.
'
' @return A reference to the single SecurityElementStatic instance.
'
Public Function SecurityElement() As SecurityElementStatic
    Set SecurityElement = modStaticClasses.SecurityElement
End Function

''
' Provides a method to access the BigInteger static methods.
'
' @return The BigInteger static methods.
'
Public Function BigInteger() As BigIntegerStatic
    Set BigInteger = modStaticClasses.BigInteger
End Function

Public Function Thread() As ThreadStatic
    Set Thread = modStaticClasses.Thread
End Function

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
' accessing static methods, a Property Get of the same name as the class that is
' supposed to contain the static methods is defined in this class. This class
' has instancing of <i>GlobalMultiUse</i>. This means that just the name of these
' Property Gets can be typed directly without having to instantiate the <b>PublicProperty Gets</b>
' class. By just typing in any of these Property Get names, the underlying class which
' contains the static methods is returned, and those methods are then available
' for calling. This gives the illusion of static methods being called on a class.
' <p>When typing in one of these Property Get names, it should be followed immediately
' by a period to allow access to the static methods of that class.</p>
' <pre>
' '' Print the name of the current user
' Debug.Print Environment.UserName
' </pre>
' This small example simply refers to the <i>Environment</i> Property Get in this
' class, then directly calls a method on the returned object from the
' <i>Environment</i> Property Get.
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
Public Property Get Cor() As Constructors
    Set Cor = Statics.Cor
End Property

''
' Returns an object used to access the <b>CorArray</b> methods
' without instantiating the <b>CorArray</b> class.
'
' @return A reference to the <b>CorArray</b> object used to
' perform Property Gets on arrays.
' @see CorArray
'
Public Property Get CorArray() As CorArray
    Set CorArray = Statics.CorArray
End Property

''
' Returns an object used to access the <b>cString</b> methods
' without instantiating the <b>cString</b> class.
'
' @return A reference to the <b>cString</b> object used to
' perform Property Gets on Strings.
' @see CorString
'
Public Property Get CorString() As CorString
    Set CorString = Statics.CorString
End Property

''
' Returns an object used to access the <b>Comparer</b> static Property Gets.
'
' @return A reference to the object used to access the <b>Comparer</b> static Property Gets.
' @see Comparer
' @see ComparerStatic
'
Public Property Get Comparer() As ComparerStatic
    Set Comparer = Statics.Comparer
End Property

''
' Returns a reference used to access the methods of the <b>Environment</b> Property Gets.
'
' @return A reference that can be used to access methods of the <b>Environment</b> class.
' @see Environment
'
Public Property Get Environment() As Environment
    Set Environment = Statics.Environment
End Property

''
' Returns a reference used to access static methods of the <b>BitArray</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitArray</b> class.
' @see BitArray
' @see BitArrayStatic
Public Property Get BitArray() As BitArrayStatic
    Set BitArray = Statics.BitArray
End Property

''
' Returns a reference used to access static methods of the <b>Buffer</b> class.
'
' @return A reference that can be used to access static methods of the <b>Buffer</b> class.
' @see Buffer
'
Public Property Get Buffer() As Buffer
    Set Buffer = Statics.Buffer
End Property

''
' Returns a reference used to access static methods of the <b>NumberFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>NumberFormatInfo</b> class.
' @see NumberFormatInfo
' @see NumberFormatInfoStatic
'
Public Property Get NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = Statics.NumberFormatInfo
End Property

''
' Returns a reference used to access static methods of the <b>BitConverter</b> class.
'
' @return A reference that can be used to access static methods of the <b>BitConverter</b> class.
' @see BitConverter
'
Public Property Get BitConverter() As BitConverter
    Set BitConverter = Statics.BitConverter
End Property

''
' Returns a reference used to access static methods of the <b>Version</b> class.
'
' @return A reference that can be used to access static methods of the <b>Version</b> class.
' @see Version
' @see VersionStatic
'
Public Property Get Version() As VersionStatic
    Set Version = Statics.Version
End Property

''
' Returns a reference used to access static methods of the <b>TimeSpan</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeSpan</b> class.
' @see TimeSpan
' @see TimeSpanStatic
'
Public Property Get TimeSpan() As TimeSpanStatic
    Set TimeSpan = Statics.TimeSpan
End Property

''
' Returns a reference used to access static methods of the <b>cDateTime</b> class.
'
' @return A reference that can be used to access static methods of the <b>cDateTime</b> class.
' @see cDateTime
' @see cDateTimeStatic
'
Public Property Get cDateTime() As cDateTimeStatic
    Set cDateTime = Statics.cDateTime
End Property

''
' Returns a reference used to access static methods of the <b>DateTimeFormatInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>DateTimeFormatInfo</b> class.
' @see DateTimeFormatInfo
' @see DateTimeFormatInfoStatic
'
Public Property Get DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = Statics.DateTimeFormatInfo
End Property

''
' Returns a reference used to access static methods of the <b>CultureInfo</b> class.
'
' @return A reference that can be used to access static methods of the <b>CultureInfo</b> class.
' @see CultureInfo
' @see CultureInfoStatic
'
Public Property Get CultureInfo() As CultureInfoStatic
    Set CultureInfo = Statics.CultureInfo
End Property

''
' Returns a reference used to access static methods of the <b>TimeZone</b> class.
'
' @return A reference that can be used to access static methods of the <b>TimeZone</b> class.
' @see TimeZone
' @see TimeZoneStatic
'
Public Property Get TimeZone() As TimeZoneStatic
    Set TimeZone = Statics.TimeZone
End Property

''
' Returns a reference used to access static methods of the <b>Path</b> class.
'
' @return A reference that can be used to access static methods of the <b>Path</b> class.
' @see Path
'
Public Property Get Path() As Path
    Set Path = Statics.Path
End Property

''
' Returns a reference used to access static methods of the <b>Encoding</b> class.
'
' @return A reference that can be used to access static methods of the <b>Encoding</b> class.
' @see Encoding
' @see EncodingStatic
'
Public Property Get Encoding() As EncodingStatic
    Set Encoding = Statics.Encoding
End Property

''
' Returns a reference used to access static methods of the <b>TextReader</b> class.
'
' @return A reference that can be used to access static methods of the <b>TextReader</b> class.
' @see TextReader
' @see TextReaderStatic
' @see StreamReader
'
Public Property Get TextReader() As TextReaderStatic
    Set TextReader = Statics.TextReader
End Property

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
Public Property Get Directory() As Directory
    Set Directory = Statics.Directory
End Property

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
Public Property Get File() As File
    Set File = Statics.File
End Property

''
' Returns a reference used to access static methods of the <b>Stream</b> class.
'
' @return A reference that can be used to access static methods of the <b>Stream</b> class.
' @see Stream
' @see StreamStatic
'
Public Property Get Stream() As StreamStatic
    Set Stream = Statics.Stream
End Property

''
' Returns a reference used to access static methods of the <b>GregorianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>GregorianCalendar</b> class.
' @see GregorianCalendar
' @see GregorianCalendarStatic
'
Public Property Get GregorianCalendar() As GregorianCalendarStatic
    Set GregorianCalendar = Statics.GregorianCalendar
End Property

''
' Returns a reference used to access static methods of the <b>JulianCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>JulianCalendar</b> class.
' @see JulianCalendar
' @see JulianCalendarStatic
'
Public Property Get JulianCalendar() As JulianCalendarStatic
    Set JulianCalendar = Statics.JulianCalendar
End Property

''
' Returns a reference used to access static methods of the <b>HebrewCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HebrewCalendar</b> class.
' @see HebrewCalendar
' @see HebrewCalendarStatic
'
Public Property Get HebrewCalendar() As HebrewCalendarStatic
    Set HebrewCalendar = Statics.HebrewCalendar
End Property

''
' Returns a reference used to access static methods of the <b>KoreanCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>KoreanCalendar</b> class.
' @see KoreanCalendar
' @see KoreanCalendarStatic
'
Public Property Get KoreanCalendar() As KoreanCalendarStatic
    Set KoreanCalendar = Statics.KoreanCalendar
End Property

''
' Returns a reference used to access static methods of the <b>HijriCalendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>HijriCalendar</b> class.
' @see HijriCalendar
' @see HijriCalendarStatic
'
Public Property Get HijriCalendar() As HijriCalendarStatic
    Set HijriCalendar = Statics.HijriCalendar
End Property

''
' Returns a reference used to access static methods of the <b>Calendar</b> class.
'
' @return A reference that can be used to access static methods of the <b>Calendar</b> class.
' @see Calendar
' @see CalendarStatic
'
Public Property Get Calendar() As CalendarStatic
    Set Calendar = Statics.Calendar
End Property

''
' Returns a reference used to access static methods of the <b>Console</b> class.
'
' @return A reference that can be used to access static methods of the <b>Console</b> class.
' @see Console
'
Public Property Get Console() As Console
    Set Console = Statics.Console
End Property

''
' Returns a reference used to access static methods of the <b>ArrayList</b> class.
'
' @return A reference that can be used to access static methods of the <b>ArrayList</b> class.
' @see ArrayList
' @see ArrayListStatic
'
Public Property Get ArrayList() As ArrayListStatic
    Set ArrayList = Statics.ArrayList
End Property

''
' Returns a reference used to access static methods of the <b>Registry</b> class.
'
' @return A reference that can be used to access static methods of the <b>Registry</b> class.
' @see Registry
' @see RegistryKey
'
Public Property Get Registry() As Registry
    Set Registry = Statics.Registry
End Property

''
' Returns a reference used to access static methods of the <b>RegistryKey</b> class.
'
' @return A reference that can be used to access static methods of the <b>RegistryKey</b> class.
' @see RegistryKey
' @see RegistryKeyStatic
' @see Registry
'
Public Property Get RegistryKey() As RegistryKeyStatic
    Set RegistryKey = Statics.RegistryKey
End Property

''
' Provides static methods for creating new <b>Guid</b> objects.
'
' @return A reference that can be used to access the static methods of the <b>Guid</b> class.
' @see GuidStatic
' @see Guid
'
Public Property Get Guid() As GuidStatic
    Set Guid = Statics.Guid
End Property

''
' Provides methods for converting values.
'
' @return A reference to the method provider.
' @see Convert
'
Public Property Get Convert() As Convert
    Set Convert = Statics.Convert
End Property

''
' Provides methods for creating custom resoure managers.
'
' @return A reference to the method provider.
' @see ResourceManagerStatic
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
'
Public Property Get ResourceManager() As ResourceManagerStatic
    Set ResourceManager = Statics.ResourceManager
End Property

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
Public Property Get DriveInfo() As DriveInfoStatic
    Set DriveInfo = Statics.DriveInfo
End Property

''
' Provides default methods for classes that implement the IObject interface.
'
' @return A reference to default implementations of the IObject interface.
'
Public Property Get Object() As ObjectStatic
    Set Object = Statics.Object
End Property

''
' Returns a reference used to access the DES static methods.
'
' @return A reference to the single DESStatic instance.
'
Public Property Get DES() As DESStatic
    Set DES = Statics.DES
End Property

''
' Returns a reference used to access the TripleDES static methods.
'
' @return A reference to the single TripleDESStatic instance.
'
Public Property Get TripleDES() As TripleDESStatic
    Set TripleDES = Statics.TripleDES
End Property

''
' Returns a reference used to access the RC2 static methods.
'
' @return A reference to the single RC2Static instance.
'
Public Property Get RC2() As RC2Static
    Set RC2 = Statics.RC2
End Property

''
' Returns a reference used to access the Rijndael static methods.
'
' @return A reference to the single RijndaelStatic instance.
'
Public Property Get Rijndael() As RijndaelStatic
    Set Rijndael = Statics.Rijndael
End Property

''
' Returns a reference used to access the CryptoConfig static methods.
'
' @return A reference to the single CryptoConfig instance.
'
Public Property Get CryptoConfig() As CryptoConfig
    Set CryptoConfig = Statics.CryptoConfig
End Property

''
' Returns a reference used to access the SHA1 static methods.
'
' @return A reference to the single SHA1Static instance.
'
Public Property Get SHA1() As SHA1Static
    Set SHA1 = Statics.SHA1
End Property

''
' Returns a reference used to access the StopWatch static methods.
'
' @return A reference to the single StopWatchStatic instance.
'
Public Property Get StopWatch() As StopWatchStatic
    Set StopWatch = Statics.StopWatch
End Property

''
' Returns a reference used to access the SHA256 static methods.
'
' @return A reference to the single SHA256Static instance.
'
Public Property Get SHA256() As SHA256Static
    Set SHA256 = Statics.SHA256
End Property

''
' Returns a reference used to access the SHA512 static methods.
'
' @return A reference to the single SHA512Static instance.
'
Public Property Get SHA512() As SHA512Static
    Set SHA512 = Statics.SHA512
End Property

''
' Returns a reference used to access the SHA384 static methods.
'
' @return A reference to the single SHA384Static instance.
'
Public Property Get SHA384() As SHA384Static
    Set SHA384 = Statics.SHA384
End Property

''
' Returns a reference used to access the MACTripleDES static methods.
'
' @return A reference to the single MACTripleDESStatic instance.
'
Public Property Get MACTripleDES() As MACTripleDESStatic
    Set MACTripleDES = Statics.MACTripleDES
End Property

''
' Returns a reference used to access the HMAC static methods.
'
' @return A reference to the single HMACStatic instance.
'
Public Property Get HMAC() As HMACStatic
    Set HMAC = Statics.HMAC
End Property

''
' Returns a reference used to access the RSA static methods.
'
' @return A reference to the single RSAStatic instance.
'
Public Property Get RSA() As RSAStatic
    Set RSA = Statics.RSA
End Property

''
' Returns a reference used to access the SecurityElement static methods.
'
' @return A reference to the single SecurityElementStatic instance.
'
Public Property Get SecurityElement() As SecurityElementStatic
    Set SecurityElement = Statics.SecurityElement
End Property

''
' Provides a method to access the BigInteger static methods.
'
' @return The BigInteger static methods.
'
Public Property Get BigInteger() As BigIntegerStatic
    Set BigInteger = Statics.BigInteger
End Property

Public Property Get Thread() As ThreadStatic
    Set Thread = Statics.Thread
End Property

Public Property Get EqualityComparer() As IEqualityComparer
    Set EqualityComparer = Statics.EqualityComparer
End Property

Public Property Get MyBase() As ObjectBase
    Set MyBase = Statics.MyBase
End Property

Public Property Get Char() As Char
    Set Char = Statics.Char
End Property

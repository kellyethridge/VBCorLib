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
' Module: StaticClasses
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
' <h4>Example</h4>
' <pre>
' ''' Print the name of the current user
' Debug.Print Environment.UserName
' </pre>
' This example simply refers to the <a href="StaticClasses.Get.Environment.html">Environment Property Get</a> in this
' class, then directly calls a method on the returned object from the <a href="StaticClasses.Get.Environment.html">Environment Property Get</a>.
' In this case it returns a reference to the <a href="Environment.html">Environment</a> class.
'@Folder("CorLib.System")
Option Explicit

''
' Represents the sign values of <a href="BigInteger.html">BigInteger</a> values.
'
' @param Negative The value is a negative number.
' @param Zero The value is zero.
' @param Positive The value is a positive number.
Public Enum Sign
    Negative = -1
    Zero = 0
    Positive = 1
End Enum


''
' Returns a object containing all of the parameterized constructors available within the VBCorLib library.
'
' @return A object containing all of the constructors in VBCorLib.
' @see Constructors
Public Property Get Cor() As Constructors
    Set Cor = Statics.Cor
End Property

''
' Gets an object used to access the <a href="CorArray.html">CorArray</a> static methods.
'
' @return A reference to a <a href="CorArray.html">CorArray</a> object.
' @see CorArray
Public Property Get CorArray() As CorArray
    Set CorArray = Statics.CorArray
End Property

''
' Gets an object used to access the <a href="CorString.html">CorString</a> static methods.
'
' @return A reference to a <a href="CorString.html">CorString</a> object.
' @see CorString
Public Property Get CorString() As CorString
    Set CorString = Statics.CorString
End Property

''
' Gets an object used to access the <a href="CorMath.html">CorMath</a> static methods.
'
' @return A reference to a <a href="CorMath.html">CorMath</a> object.
Public Property Get CorMath() As CorMath
    Set CorMath = Statics.CorMath
End Property

''
' Gets an object used to access the <a href="Comparer.html">Comparer</a> static methods.
'
' @return A reference to a <a href="ComparerStatic.html">ComparerStatic</a> object.
' @see Comparer
' @see ComparerStatic
Public Property Get Comparer() As ComparerStatic
    Set Comparer = Statics.Comparer
End Property

''
' Gets an object used to access the <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> static methods.
'
' @return A reference to a <a href="CaseInsensitiveComparerStatic.html">CaseInsensitiveComparerStatic</a> object.
' @see CaseInsensitiveComparer
' @see CaseInsensitiveComparerStatic
Public Property Get CaseInsensitiveComparer() As CaseInsensitiveComparerStatic
    Set CaseInsensitiveComparer = Statics.CaseInsensitiveComparer
End Property

''
' Gets an object used to access the <a href="Environment.html">Environment</a> static methods.
'
' @return A reference to an <a href="Environment.html">Environment</a> object.
' @see Environment
Public Property Get Environment() As Environment
    Set Environment = Statics.Environment
End Property

''
' Gets an object used to access the <a href="BitArray.html">BitArray</a> static methods.
'
' @return A reference to a <a href="BitArrayStatic.html">BitArrayStatic</a> object.
' @see BitArray
' @see BitArrayStatic
Public Property Get BitArray() As BitArrayStatic
    Set BitArray = Statics.BitArray
End Property

''
' Gets an object used to access the <a href="Buffer.html">Buffer</a> static methods.
'
' @return A reference to a <a href="Buffer.html">Buffer</a> object.
' @see Buffer
Public Property Get Buffer() As Buffer
    Set Buffer = Statics.Buffer
End Property

''
' Gets an object used to access the <a href="NumberFormatInfo.html">NumberFormatInfo</a> static methods.
'
' @return A reference to a <a href="NumberFormatInfoStatic.html">NumberFormatInfoStatic</a> object.
' @see NumberFormatInfo
' @see NumberFormatInfoStatic
Public Property Get NumberFormatInfo() As NumberFormatInfoStatic
    Set NumberFormatInfo = Statics.NumberFormatInfo
End Property

''
' Gets an object used to access the <a href="BitConverter.html">BitConverter</a> static methods.
'
' @return A reference to a <a href="BitConverter.html">BitConverter</a> object.
' @see BitConverter
Public Property Get BitConverter() As BitConverter
    Set BitConverter = Statics.BitConverter
End Property

''
' Gets an object used to access the <a href="Version.html">Version</a> static methods.
'
' @return A reference to a <a href="VersionStatic.html">VersionStatic</a> object.
' @see Version
' @see VersionStatic
Public Property Get Version() As VersionStatic
    Set Version = Statics.Version
End Property

''
' Gets an object used to access the <a href="TimeSpan.html">TimeSpan</a> static methods.
'
' @return A reference to a <a href="TimeSpanStatic.html">TimeSpanStatic</a> object.
' @see TimeSpan
' @see TimeSpanStatic
Public Property Get TimeSpan() As TimeSpanStatic
    Set TimeSpan = Statics.TimeSpan
End Property

''
' Gets an object used to access the <a href="CorDateTime.html">CorDateTime</a> static methods.
'
' @return A reference to a <a href="CorDateTimeStatic.html">CorDateTimeStatic</a> object.
' @see CorDateTime
' @see CorDateTimeStatic
Public Property Get CorDateTime() As CorDateTimeStatic
    Set CorDateTime = Statics.CorDateTime
End Property

''
' Gets an object used to access the <a href="DateTimeFormatInfo.html">DateTimeFormatInfo</a> static methods.
'
' @return A reference to a <a href="DateTimeFormatInfoStatic.html">DateTimeFormatInfoStatic</a> object.
' @see DateTimeFormatInfo
' @see DateTimeFormatInfoStatic
Public Property Get DateTimeFormatInfo() As DateTimeFormatInfoStatic
    Set DateTimeFormatInfo = Statics.DateTimeFormatInfo
End Property

''
' Gets an object used to access the <a href="CultureInfo.html">CultureInfo</a> static methods.
'
' @return A reference to a <a href="CultureInfoStatic.html">CultureInfoStatic</a> object.
' @see CultureInfo
' @see CultureInfoStatic
Public Property Get CultureInfo() As CultureInfoStatic
    Set CultureInfo = Statics.CultureInfo
End Property

''
' Gets an object used to access the <a href="TimeZone.html">TimeZone</a> static methods.
'
' @return A reference to a <a href="TimeZoneStatic.html">TimeZoneStatic</a> object.
' @see TimeZone
' @see TimeZoneStatic
Public Property Get TimeZone() As TimeZoneStatic
    Set TimeZone = Statics.TimeZone
End Property

''
' Gets an object used to access the <a href="Path.html">Path</a> static methods.
'
' @return A reference to a <a href="Path.html">Path</a> object.
' @see Path
Public Property Get Path() As Path
    Set Path = Statics.Path
End Property

''
' Gets an object used to access the <a href="Encoding.html">Encoding</a> static methods.
'
' @return A reference to an <a href="EncodingStatic.html">EncodingStatic</a> object.
' @see Encoding
' @see EncodingStatic
Public Property Get Encoding() As EncodingStatic
    Set Encoding = Statics.Encoding
End Property

''
' Gets an object used to access the <a href="TextReader.html">TextReader</a> static methods.
'
' @return A reference to a <a href="TextReaderStatic.html">TextReaderStatic</a> object.
' @see TextReader
' @see TextReaderStatic
' @see StreamReader
Public Property Get TextReader() As TextReaderStatic
    Set TextReader = Statics.TextReader
End Property

''
' Gets an object used to access the <a href="Directory.html">Directory</a> static methods.
'
' @return A reference to a <a href="Directory.html">Directory</a> object.
' @see Directory
' @see DirectoryInfo
' @see File
' @see FileInfo
' @see DriveInfo
' @see DriveInfoStatic
Public Property Get Directory() As Directory
    Set Directory = Statics.Directory
End Property

''
' Gets an object used to access the <a href="File.html">File</a> static methods.
'
' @return A reference to a <a href="File.html">File</a> object.
' @see File
' @see FileInfo
' @see Directory
' @see DirectoryInfo
' @see DriveInfo
' @see DriveInfoStatic
Public Property Get File() As File
    Set File = Statics.File
End Property

''
' Gets an object used to access the <a href="Stream.html">Stream</a> static methods.
'
' @return A reference to a <a href="StreamStatic.html">StreamStatic</a> object.
' @see Stream
' @see StreamStatic
Public Property Get Stream() As StreamStatic
    Set Stream = Statics.Stream
End Property

''
' Gets an object used to access the <a href="GregorianCalendar.html">GregorianCalendar</a> static methods.
'
' @return A reference to a <a href="GregorianCalendarStatic.html">GregorianCalendarStatic</a> object.
' @see GregorianCalendar
' @see GregorianCalendarStatic
Public Property Get GregorianCalendar() As GregorianCalendarStatic
    Set GregorianCalendar = Statics.GregorianCalendar
End Property

''
' Gets an object used to access the <a href="JulianCalendar.html">JulianCalendar</a> static methods.
'
' @return A reference to a <a href="JulianCalendarStatic.html">JulianCalendarStatic</a> object.
' @see JulianCalendar
' @see JulianCalendarStatic
Public Property Get JulianCalendar() As JulianCalendarStatic
    Set JulianCalendar = Statics.JulianCalendar
End Property

''
' Gets an object used to access the <a href="HewbrewCalendar.html">HebrewCalendar</a> static methods.
'
' @return A reference to a <a href="HebrewCalendarStatic.html">HebrewCalendarStatic</a> object.
' @see HebrewCalendar
' @see HebrewCalendarStatic
Public Property Get HebrewCalendar() As HebrewCalendarStatic
    Set HebrewCalendar = Statics.HebrewCalendar
End Property

''
' Gets an object used to access the <a href="KoreanCalendar.html">KoreanCalendar</a> static methods.
'
' @return A reference to a <a href="KoreanCalendarStatic.html">KoreanCalendarStatic</a> object.
' @see KoreanCalendar
' @see KoreanCalendarStatic
Public Property Get KoreanCalendar() As KoreanCalendarStatic
    Set KoreanCalendar = Statics.KoreanCalendar
End Property

''
' Gets an object used to access the <a href="HijriCalendar.html">HijriCalendar</a> static methods.
'
' @return A reference to a <a href="HijriCalendarStatic.html">HijriCalendarStatic</a> object.
' @see HijriCalendar
' @see HijriCalendarStatic
Public Property Get HijriCalendar() As HijriCalendarStatic
    Set HijriCalendar = Statics.HijriCalendar
End Property

''
' Gets an object used to access the <a href="Calendar.html">Calendar</a> static methods.
'
' @return A reference to a <a href="CalendarStatic.html">CalendarStatic</a> object.
' @see Calendar
' @see CalendarStatic
Public Property Get Calendar() As CalendarStatic
    Set Calendar = Statics.Calendar
End Property

''
' Gets an object used to access the <a href="Console.html">Console</a> static methods.
'
' @return A reference to a <a href="Console.html">Console</a> object.
' @see Console
Public Property Get Console() As Console
    Set Console = Statics.Console
End Property

''
' Gets an object used to access the <a href="ArrayList.html">ArrayList</a> static methods.
'
' @return A reference to an <a href="ArrayListStatic.html">ArrayListStatic</a> object.
' @see ArrayList
' @see ArrayListStatic
Public Property Get ArrayList() As ArrayListStatic
    Set ArrayList = Statics.ArrayList
End Property

''
' Gets an object used to access the <a href="Registry.html">Registry</a> static methods.
'
' @return A reference to a <a href="Registry.html">Registry</a> object.
' @see Registry
' @see RegistryKey
'
Public Property Get Registry() As Registry
    Set Registry = Statics.Registry
End Property

''
' Gets an object used to access the <a href="RegistryKey.html">RegistryKey</a> static methods.
'
' @return A reference to a <a href="RegistryKeyStatic.html">RegistryKeyStatic</a> object.
' @see RegistryKey
' @see RegistryKeyStatic
' @see Registry
Public Property Get RegistryKey() As RegistryKeyStatic
    Set RegistryKey = Statics.RegistryKey
End Property

''
' Gets an object used to access the <a href="Guid.html">Guid</a> static methods.
'
' @return A reference to a <a href="GuidStatic.html">GuidStatic</a> object.
' @see GuidStatic
' @see Guid
Public Property Get Guid() As GuidStatic
    Set Guid = Statics.Guid
End Property

''
' Gets an object used to access the <a href="Convert.html">Convert</a> static methods.
'
' @return A reference to a <a href="Convert.html">Convert</a> object.
' @see Convert
Public Property Get Convert() As Convert
    Set Convert = Statics.Convert
End Property

''
' Gets an object used to access the <a href="ResourceManager.html">ResourceManager</a> static methods.
'
' @return A reference to a <a href="ResourceManagerStatic.html">ResourceManagerStatic</a> object.
' @see ResourceManagerStatic
' @see ResourceManager
' @see ResourceSet
' @see ResourceReader
Public Property Get ResourceManager() As ResourceManagerStatic
    Set ResourceManager = Statics.ResourceManager
End Property

''
' Gets an object used to access the <a href="DriveInfo.html">DriveInfo</a> static methods.
'
' @return A reference to a <a href="DriveInfoStatic.html">DriveInfoStatic</a> object.
' @see DriveInfo
' @see DirectoryInfo
' @see Directory
' @see FileInfo
' @see File
Public Property Get DriveInfo() As DriveInfoStatic
    Set DriveInfo = Statics.DriveInfo
End Property

''
' Gets an object used to access the <a href="ObjectStatic.html">ObjectStatic</a> static methods.
'
' @return A reference to <a href="ObjectStatic.html">ObjectStatic</a> object.
Public Property Get Object() As ObjectStatic
    Set Object = Statics.Object
End Property

''
' Gets an object used to access the <a href="DES.html">DES</a> static methods.
'
' @return A reference to a <a href="DESStatic.html">DESStatic</a> object.
Public Property Get DES() As DESStatic
    Set DES = Statics.DES
End Property

''
' Gets an object used to access the <a href="TripleDES.html">TripleDES</a> static methods.
'
' @return A reference to a <a href="TripleDESStatic.html">TripleDESStatic</a> object.
Public Property Get TripleDES() As TripleDESStatic
    Set TripleDES = Statics.TripleDES
End Property

''
' Gets an object used to access the <a href="RC2.html">RC2</a> static methods.
'
' @return A reference to an <a href="RC2Static.html">RC2Static</a> object.
Public Property Get RC2() As RC2Static
    Set RC2 = Statics.RC2
End Property

''
' Gets an object used to access the <a href="Rijndael.html">Rijndael</a> static methods.
'
' @return A reference to a <a href="RijndaelStatic.html">RijndailStatic</a> object.
Public Property Get Rijndael() As RijndaelStatic
    Set Rijndael = Statics.Rijndael
End Property

''
' Gets an object used to access the <a href="CryptoConfig.html">CryptoConfig</a> static methods.
'
' @return A reference to a <a href="CryptoConfig.html">CryptoConfig</a> object.
Public Property Get CryptoConfig() As CryptoConfig
    Set CryptoConfig = Statics.CryptoConfig
End Property

''
' Gets an object used to access the <a href="StopWatch.html">StopWatch</a> static methods.
'
' @return A reference to a <a href="StopWatchStatic.html">StopWatchStatic</a> object.
Public Property Get StopWatch() As StopWatchStatic
    Set StopWatch = Statics.StopWatch
End Property

''
' Gets an object used to access the <a href="MD5.html">MD5</a> static methods.
'
' @return A reference to an <a href="MD5Static.html">MD5Static</a> object.
Public Property Get MD5() As MD5Static
    Set MD5 = Statics.MD5
End Property

''
' Gets an object used to access the <a href="SHA1.html">SHA1</a> static methods.
'
' @return A reference to an <a href="SHA1Static.html">SHA1Static</a> object.
Public Property Get SHA1() As SHA1Static
    Set SHA1 = Statics.SHA1
End Property

''
' Gets an object used to access the <a href="SHA256.html">SHA256</a> static methods.
'
' @return A reference to an <a href="SHA256Static.html">SHA256Static</a> object.
Public Property Get SHA256() As SHA256Static
    Set SHA256 = Statics.SHA256
End Property

''
' Gets an object used to access the <a href="SHA512.html">SHA512</a> static methods.
'
' @return A reference to an <a href="SHA512Static.html">SHA512Static</a> object.
Public Property Get SHA512() As SHA512Static
    Set SHA512 = Statics.SHA512
End Property

''
' Gets an object used to access the <a href="SHA384.html">SHA384</a> static methods.
'
' @return A reference to an <a href="SHA384Static.html">SHA384Static</a> object.
Public Property Get SHA384() As SHA384Static
    Set SHA384 = Statics.SHA384
End Property

''
' Gets an object used to access the <a href="MACTripleDES.html">MACTripleDES</a> static methods.
'
' @return A reference to a <a href="MACTripleDESStatic.html">MACTripleDESStatic</a> object.
Public Property Get MACTripleDES() As MACTripleDESStatic
    Set MACTripleDES = Statics.MACTripleDES
End Property

''
' Gets an object used to access the <a href="HMAC.html">HMAC</a> static methods.
'
' @return A reference to an <a h ref="HMACStatic.html">HMACStatic</a> object.
Public Property Get HMAC() As HMACStatic
    Set HMAC = Statics.HMAC
End Property

''
' Gets an object used to access the <a href="RSA.html">RSA</a> static methods.
'
' @return A reference to a <a href="RSAStatic.html">RSAStatic</a> object.
Public Property Get RSA() As RSAStatic
    Set RSA = Statics.RSA
End Property

''
' Gets an object used to access the <a href="SecurityElement.html">SecurityElement</a> static methods.
'
' @return A reference to a <a href="SecurityElementStatic.html">SecurityElementStatic</a> object.
Public Property Get SecurityElement() As SecurityElementStatic
    Set SecurityElement = Statics.SecurityElement
End Property

''
' Gets an object used to access the <a href="BigInteger.html">BigInteger</a> static methods.
'
' @return A reference to a <a href="BigIntegerStatic.html">BigIntegerStatic</a> object.
Public Property Get BigInteger() As BigIntegerStatic
    Set BigInteger = Statics.BigInteger
End Property

''
' Gets an object used to access the <a href="Thread.html">Thread</a> static methods.
'
' @return A reference to a <a href="ThreadStatic.html">ThreadStatic</a> object.
Public Property Get Thread() As ThreadStatic
    Set Thread = Statics.Thread
End Property

''
' Gets a reference used to access the <a href="EqualityComparer.html">EqualityComparer</a> static methods.
'
' @return A reference to an <a href="EqualityComparerStatic.html">EqualityComparerStatic</a> object.
Public Property Get EqualityComparer() As EqualityComparerStatic
    Set EqualityComparer = Statics.EqualityComparer
End Property

''
' Gets a reference used to access the <a href="ObjectBase.html">ObjectBase</a> static methods.
'
' @return A reference to an <a href="ObjectBase.html">ObjectBase</a> object.
' @remarks The <a href="ObjectBase.html">ObjectBase</a> object provides default
' implementations of the methods in the <a href="IObject.html">IObject</a> interface.
Public Property Get MyBase() As ObjectBase
    Set MyBase = Statics.MyBase
End Property

''
' Gets a reference used to access the <a href="Char.html">Char</a> static methods.
'
' @return A reference to a <a href="Char.html">Char</a> object.
Public Property Get Char() As Char
    Set Char = Statics.Char
End Property

''
' Gets a reference used to access the <a href="StringComparer.html">StringComparer</a> static methods.
'
' @return A reference to a <a href="StringComparerStatic.html">StringComparerStatic</a> object.
Public Property Get StringComparer() As StringComparerStatic
    Set StringComparer = Statics.StringComparer
End Property

''
' Gets a reference used to access the <a href="RSACryptoServiceProvider.html">RSACryptoServiceProvider</a> static methods.
'
' @return A reference to an <a href="RSACryptoServiceProviderStatic.html">RSACryptoServiceProviderStatic</a> object.
Public Property Get RSACryptoServiceProvider() As RSACryptoServiceProviderStatic
    Set RSACryptoServiceProvider = Statics.RSACryptoServiceProvider
End Property

''
' Gets a reference used to access the <a href="DecoderFallback.html">DecoderFallback</a> static methods.
'
' @return A reference to a <a href="DecoderFallbackStatic.html">DecoderFallbackStatic</a> object.
Public Property Get DecoderFallback() As DecoderFallbackStatic
    Set DecoderFallback = Statics.DecoderFallback
End Property

''
' Gets a reference used to access the <a href="EncoderFallback.html">EncoderFallback</a> static methods.
'
' @return A reference to a <a href="EncoderFallbackStatic.html">EncoderFallbackStatic</a> object.
Public Property Get EncoderFallback() As EncoderFallbackStatic
    Set EncoderFallback = Statics.EncoderFallback
End Property

''
' Gets a reference used to access the <a href="MemoryMappedFile.html">MemoryMappedFile</a> static methods.
'
' @return A reference to a <a href="MemoryMappedFileStatic.html">MemoryMappedFileStatic</a> object.
Public Property Get MemoryMappedFile() As MemoryMappedFileStatic
    Set MemoryMappedFile = Statics.MemoryMappedFile
End Property

''
' Gets a reference used to access the <a href="Int64Static.Int64.html">Int64</a> static methods.
'
' @return A reference to an <a href="Int64Static.html">Int64Static</a> object.
Public Property Get Int64() As Int64Static
    Set Int64 = Statics.Int64
End Property

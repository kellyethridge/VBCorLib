VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constructors"
Attribute VB_GlobalNameSpace = True
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
' Module: Constructors
'

''
' Functions used to create fully inititalized objects.
'
' @remarks This class is defined as <b>Global-Multiuse</b>. All of the functions can be called
' directly without needing to create a <b>Constructors</b> object.
' <p>The <a href="StaticClasses.Get.Cor.html">Cor</a> accessor can be used to easily find and select the
' desired object to create. Using the <a href="StaticClasses.Get.Cor.html">Cor</a> accessor provides
' IntelliSense to the available constructors, making object creation easier.</p>
' @include "..\DocIncludes\Constructors.html"
Option Explicit

Private mBigIntegerGuid As VBGUID


''
' Creates a <i>For..Each</i> compatible wrapper around an <a href="IEnumerator.html">IEnumerator</a> interface.
'
' @param obj The <a href="IEnumerator.html">IEnumerator</a> object that is to be called during For..Each.
' @return A <i>For..Each</i> compatible enumerator.
' @remarks VBCorLib allows for the building of custom enumerators to be used
' in a <i>For..Each</i> loop context.
' <p>To create a custom enumerator please refer to the <a href="IEnumerator.html">IEnumerator</a> interface.</p>
' <p>In order for Visual Basic to utilize a custom <a href="IEnumerator.html">IEnumerator</a> object, it must be
' wrapped in an object that implements the <b>IEnumVariant</b> interface. This is the
' only interface that Visual Basic can use in a <i>For..Each</i> context. By passing in the
' custom enumerator into the <b>CreateEnumerator</b> function, a wrapper is applied and returned that
' is then returned in the typical <b>NewEnum</b> function of a class.
' @see IEnumerator
' @include "..\DocIncludes\Constructors.CreateEnumerator.html"
Public Function CreateEnumerator(ByVal Obj As IEnumerator) As IUnknown
    Set CreateEnumerator = Enumeration.CreateEnumerator(Obj)
End Function

''
' Creates a new <a href="Exception.html">Exception</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param ErrorNumber Represents a VBA specific error number.
' @param InnerException An exception that caused this exception.
' @return The fully created Exception object.
' @see Exception
' @remarks The default error number is 5-Invalid Procedure Call.
Public Function NewException(ByRef Message As String, Optional ByVal ErrorNumber As Long = vbInvalidProcedureCall, Optional ByVal InnerException As Exception) As Exception
    Set NewException = New Exception
    NewException.Init Message, ErrorNumber, InnerException
End Function

''
' Creates a new <a href="SystemException.html">SystemException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="SystemException.html">SystemException</a> object.
' @see SystemException
' @see Exception
Public Function NewSystemException(ByRef Message As String, Optional ByVal InnerException As Exception) As SystemException
    Set NewSystemException = New SystemException
    NewSystemException.Init Message, InnerException
End Function

''
' Creates a new <a href="UnauthorizedAccessException.html">UnauthorizedAccessException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="UnauthorizedAccessException.html">UnauthorizedAccessException</a> object.
Public Function NewUnauthorizedAccessException(ByRef Message As String, Optional ByVal InnerException As Exception) As UnauthorizedAccessException
    Set NewUnauthorizedAccessException = New UnauthorizedAccessException
    NewUnauthorizedAccessException.Init Message, InnerException
End Function

''
' Creates a new <a href="NotSupportedException.html">NotSupportedException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="NotSupportedException.html">NotSupportedException</a> object.
Public Function NewNotSupportedException(ByRef Message As String, Optional ByVal InnerException As Exception) As NotSupportedException
    Set NewNotSupportedException = New NotSupportedException
    NewNotSupportedException.Init Message, InnerException
End Function

''
' Creates a new <a href="ArgumentException.html">ArgumentException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param ParamName The parameter that caused the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="ArgumentException.html">ArgumentException</a> object.
' @remarks <p>When throwing an <a href="ArgumentException.html">ArgumentException</a> a meaningful error message describing the invalid argument, the name
' of the argument, as well as the expected range of values for the argument should be defined.</p>
' @see ArgumentException
' @include "..\DocIncludes\Constructors.NewArgumentException.html"
Public Function NewArgumentException(ByRef Message As String, Optional ByRef ParamName As String, Optional ByVal InnerException As Exception) As ArgumentException
    Set NewArgumentException = New ArgumentException
    NewArgumentException.Init Message, ParamName, InnerException
End Function

''
' Creates a new <a href="ArgumentNullException.html">ArgumentNullException</a> object.
'
' @param ParamName The parameter that caused the exception.
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="ArgumentNullException.html">ArgumentNullException</a> object.
' @remarks <p>When throwing an <a href="ArgumentNullException.html">ArgumentNullException</a> a meaningful error message describing the invalid argument
' and the name of the argument should be defined.</p>
' @examples
' <p><pre>
' '' Throws a default ArgumentNullException object
' Throw New ArgumentNullException
'
' '' Throws an ArgumentNullException with a custom message and parameter name
' Throw Cor.NewArgumentNullException("MyParam", "A valid object is required.")
'
' '' Throws an ArgumentNullException with a custom message and error code and parameter name
' Dim ex As ArgumentNullException
' Set ex = Cor.NewArgumentNullException("MyParam", "A valid object is required.")
' ex.HResult = 123
' Throw ex
' </pre></p>
' This example shows how to throw three different variations of the same <a href="ArgumentNullException.html">ArgumentNullException</a> object.
' @see ArgumentNullException
Public Function NewArgumentNullException(Optional ByRef ParamName As String, Optional ByRef Message As String, Optional ByVal InnerException As Exception) As ArgumentNullException
    Set NewArgumentNullException = New ArgumentNullException
    NewArgumentNullException.Init ParamName, Message, InnerException
End Function

''
' Creates a new <a href="ArgumentOutOfRangeException">ArgumentOutOfRangeException</a> object.
'
' @param ParamName The parameter that caused the exception.
' @param Message A human-readable message explaining the exception.
' @param ActualValue The value of the parameter that caused the exception.
' @param InnerException An exception that caused this exception.
' @param Method Describes the type of method being executed which caused the exception.
' @return A fully created <a href="ArgumentOutOfRangeException">ArgumentOutOfRangeException</a> object.
' @see ArgumentOutOfRangeException
Public Function NewArgumentOutOfRangeException(Optional ByRef ParamName As String, Optional ByRef Message As String, Optional ByRef ActualValue As Variant, Optional ByVal InnerException As Exception, Optional Method As InvocationMethod) As ArgumentOutOfRangeException
    Set NewArgumentOutOfRangeException = New ArgumentOutOfRangeException
    NewArgumentOutOfRangeException.Init ParamName, Message, ActualValue, InnerException, Method
End Function

''
' Creates a new <a href="InvalidCaseException.html">InvalidCaseException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="InvalidCaseException.html">InvalidCastException</a> object.
Public Function NewInvalidCastException(ByRef Message As String, Optional ByVal InnerException As Exception) As InvalidCastException
    Set NewInvalidCastException = New InvalidCastException
    NewInvalidCastException.Init Message, InnerException
End Function

''
' Creates a new <a href="IndexOutOfRangeException.html">IndexOutOfRangeException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="IndexOutOfRangeException.html">IndexOutOfRangeException</a> object.
Public Function NewIndexOutOfRangeException(ByRef Message As String, Optional ByVal InnerException As Exception) As IndexOutOfRangeException
    Set NewIndexOutOfRangeException = New IndexOutOfRangeException
    NewIndexOutOfRangeException.Init Message, InnerException
End Function

''
' Creates a new <a href="RankException.html">RankException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="RankException.html">RankException</a> object.
Public Function NewRankException(ByRef Message As String, Optional ByVal InnerException As Exception) As RankException
    Set NewRankException = New RankException
    NewRankException.Init Message, InnerException
End Function

''
' Creates a new <a href="InvalidOperationException.html">InvalidOperationException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="InvalidOperationException.html">InvalidOperationException</a> object.
Public Function NewInvalidOperationException(ByRef Message As String, Optional ByVal InnerException As Exception) As InvalidOperationException
    Set NewInvalidOperationException = New InvalidOperationException
    NewInvalidOperationException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="ArrayList.html">ArrayList</a> class that contains elements copied from the specified source
' and that has the same initial capacity is at least the number of elements copied.
'
' @param Source The collection or array whose elements are copied to the new list.
' @return A new <a href="ArrayList.html">ArrayList</a> object containing the collection of values.
' @remarks The <a href="ArrayList.html">ArrayList</a> constructor accepts a <b>VBA.Collection</b>, <a href="ICollection.html">ICollection</a>, or an array as a list of
' items to initialize the new <a href="ArrayList.html">ArrayList</a> with.
' <p>The capacity of an <a href="ArrayList.html">ArrayList</a> is the number of elements that the <a href="ArrayList.html">ArrayList</a> can hold.
' As elements are added to an <a href="ArrayList.html">ArrayList</a>, the capacity is automatically increased as required by reallocating the internal array.</p>
' <p>The elements are copied onto the <a href="ArrayList.html">ArrayList</a> in the same order they are read by the For..Each statement.</p>
' <p>This constructor is an O(<i>n</i>) operation, where <i>n</i> is the number of elements in <i>Source</i>.</p>
' @see ArrayList
' @see IComparer
Public Function NewArrayList(ByRef Source As Variant) As ArrayList
    Set NewArrayList = New ArrayList
    NewArrayList.Init Source
End Function

''
' Creates a new <a href="SerializationException.html">SerializationException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="SerializationException.html">SerializationException</a> object.
Public Function NewSerializationException(ByRef Message As String, Optional ByVal InnerException As Exception) As SerializationException
    Set NewSerializationException = New SerializationException
    NewSerializationException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="Stack.html">Stack</a> class that contains elements copied from the specified collection.
'
' @param Source A collection of items copy elements from.
' @return A new <a href="Stack.html">Stack</a> object containing the collection of values.
' @remarks <i>Source</i> can be a <b>VBA.Collection</b>, an object implementing <a href="ICollection.html">ICollection</a>, or an array.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Source</i> is <b>Nothing</b> or an uninitialized array.</td>
'   </tr>
' </table>
Public Function NewStack(ByRef Source As Variant) As Stack
    Set NewStack = New Stack
    NewStack.Init Source
End Function

''
' Creates a new <a href="ArrayTypeMismatchException.html">ArrayTypeMismatchException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="ArrayTypeMismatchException.html">ArrayTypeMismatchException</a> object.
Public Function NewArrayTypeMismatchException(ByRef Message As String, Optional ByVal InnerException As Exception) As ArrayTypeMismatchException
    Set NewArrayTypeMismatchException = New ArrayTypeMismatchException
    NewArrayTypeMismatchException.Init Message, InnerException
End Function

''
' Creates a new <a href="DictionaryEntry.html">DictionaryEntry</a> object initialized with <i>Key</i> and <i>Value</i>.
'
' @param Key A value designated as the key of the pair.
' @param value A value designated as the value of the pair.
' @return A new <a href="DictionaryEntry.html">DictionaryEntry</a> containing a key/value pair.
Public Function NewDictionaryEntry(ByRef Key As Variant, ByRef Value As Variant) As DictionaryEntry
    Set NewDictionaryEntry = New DictionaryEntry
    NewDictionaryEntry.Init Key, Value
End Function

''
' Creates a new <a href="SortedList.html">SortedList</a> object with specific initialization options.
'
' @param d An object implementing <a href="IDictionary.html">IDictionary</a> used to initialize the list from.
' @param Comparer An object implementing <a href="IComparer.html">IComparer</a> used to search and sort keys in the list.
' @param Capacity An initial list size.
' @return A new <a href="SortedList.html">SortedList</a> object with the specific initializations.
Public Function NewSortedList(Optional ByVal d As IDictionary, Optional ByVal Comparer As IComparer, Optional ByVal Capacity As Long) As SortedList
    Set NewSortedList = New SortedList
    NewSortedList.Init d, Comparer, Capacity
End Function

''
' Creates a new <a href="BitArray.html">BitArray</a> object.
'
' @param Length The number of bits the object will represent.
' @param Value The value to initialize all the bits to (True = 1, False = 0).
' @return A new <a href="BitArray.html">BitArray</a> object set to the specified length with the specified value.
Public Function NewBitArray(ByVal Length As Long, Optional ByVal Value As Boolean) As BitArray
    Set NewBitArray = New BitArray
    NewBitArray.Init Length, Value
End Function

''
' Creates a new <a href="FormatException.html">FormatException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="FormatException.html">FormatException</a> object.
Public Function NewFormatException(ByRef Message As String, Optional ByVal InnerException As Exception) As FormatException
    Set NewFormatException = New FormatException
    NewFormatException.Init Message, InnerException
End Function

''
' Creates a new <a href="Random.html">Random</a> object initialized with the specified seed.
'
' @param seed The value used to initialize the random number generator with.
' @return A new <a href="Random.html">Random</a> object.
Public Function NewRandom(ByVal seed As Long) As Random
    Set NewRandom = New Random
    NewRandom.Init seed
End Function

''
' Creates a new <a href="Version.html">Version</a> object with specified version information.
'
' @param Major The number to describe an application level change.
' @param Minor The number that describes a component level change.
' @param Build The number used to describe an integration compiling of the application.
' @param Revision Minor adjustments to the application.
' @return A new <a href="Version.html">Version</a> object.
Public Function NewVersion(ByVal Major As Long, ByVal Minor As Long, Optional ByRef Build As Variant, Optional ByRef Revision As Variant) As Version
    Set NewVersion = New Version
    NewVersion.Init Major, Minor, Build, Revision
End Function

''
' Creates a new <a href="TimeSpan.html">TimeSpan</a> object set to the specified amount of time.
'
' @param Days The number of days.
' @param Hours The number of hours.
' @param Minutes The number of minutes.
' @param Seconds The number of seconds.
' @param Milliseconds The number of milliseconds.
' @return A <a href="TimeSpan.html">TimeSpan</a> with the specified time.
' @remarks The values can be both positive and negative.
Public Function NewTimeSpan(Optional ByVal Days As Long, Optional ByVal Hours As Long, Optional ByVal Minutes As Long, Optional ByVal Seconds As Long, Optional ByVal Milliseconds As Long) As TimeSpan
    Set NewTimeSpan = New TimeSpan
    NewTimeSpan.Init Days, Hours, Minutes, Seconds, Milliseconds
End Function

''
' Creates a new <a href="CorDateTime.html">CorDateTime</a> object set to the specified date.
'
' @param Year The year (1 through 9999 or the number of years in <i>Calendar</i>).
' @param Month The month (1 through 12 or the number of months in <i>Calendar</i>).
' @param Day The day (1 through the number of days in <i>Month</i>).
' @param Hour The hours (0 through 23).
' @param Minute The minutes (0 through 59).
' @param Second The seconds (0 through 59).
' @param Millisecond The millisecond (0 through 999).
' @param Calendar The calendar to use to create a <a href="CorDateTime.html">CorDateTime</a> for a specific culture.
' @param Kind Specifies whether the <a href="CorDateTime.html">CorDateTime</a> is a UTC, Local, or unspecified time type.
' @return A new <a href="CorDateTime.html">CorDateTime</a> object initialized to the specified date and time.
' @remarks <p>The <a href="GregorianCalendar.html">GregorianCalendar</a> is used if <i>Calendar</i> is not specified.</p>
Public Function NewDateTime(Optional ByVal Year As Long, Optional ByVal Month As Long, Optional ByVal Day As Long, Optional ByVal Hour As Long, Optional ByVal Minute As Long, Optional ByVal Second As Long, Optional ByVal Millisecond As Long, Optional ByVal Calendar As Calendar, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As CorDateTime
    Set NewDateTime = New CorDateTime
    NewDateTime.Init Year, Month, Day, Hour, Minute, Second, Millisecond, Calendar, Kind
End Function

''
' Creates a new <a href="IOException.html">IOException</a> object with the specified Message, HResult, and InnerException.
'
' @param Message A human-readable error message.
' @param HResult The underlying error code.
' @param ErrorCode The VB error number.
' @param InnerException The exception that caused this exception.
' @return A new <a href="IOException.html">IOException</a> object.
Public Function NewIOException(ByRef Message As String, Optional ByVal HResult As Long = COR_E_IO, Optional ByVal ErrorCode As Long = vbDeviceIOError, Optional ByVal InnerException As Exception) As IOException
    Set NewIOException = New IOException
    NewIOException.Init Message, HResult, ErrorCode, InnerException
End Function

''
' Creates a new <a href="FileNotFoundException.html">FileNotFoundException</a> with the specified <i>FileName</i>, <i>Message</i>, and <i>InnerException</i>.
'
' @param FileName The filename that caused the exception.
' @param Message A human-readable error message.
' @param InnerException The exception that caused this exception.
' @return A new <a href="FileNotFoundException.html">FileNotFoundException</a> with the specified values.
Public Function NewFileNotFoundException(Optional ByRef FileName As String, Optional ByRef Message As String, Optional ByVal InnerException As Exception) As FileNotFoundException
    Set NewFileNotFoundException = New FileNotFoundException
    NewFileNotFoundException.Init FileName, Message, InnerException
End Function

''
' Creates a new <a href="CultureInfo.html">CultureInfo</a> object for the specified culture.
'
' @param NameOrLCID The name or Locale ID of the culture to be created.
' @param UseUserOverride A flag indicating if the local machine settings should be used instead of the stored setting. This only applies to the culture of the current machine.
' @return The newly created <a href="CultureInfo.html">CultureInfo</a> object.
Public Function NewCultureInfo(ByRef NameOrLCID As Variant, Optional ByVal UseUserOverride As Boolean = True) As CultureInfo
    Set NewCultureInfo = New CultureInfo
    NewCultureInfo.Init NameOrLCID, UseUserOverride
End Function

''
' Creates a new <a href="DayLightTime.html">DayLightTime</a> object with the specified start and end times for daylight saving.
'
' @param StartTime The starting of the daylight saving period.
' @param EndTime The ending of the daylight saving period.
' @param Delta The time difference between daylight saving time and standard time.
' @return A newly created <a href="DayLightTime.html">DayLightTime</a> object.
Public Function NewDayLightTime(ByRef StartTime As Variant, ByRef EndTime As Variant, ByVal Delta As TimeSpan) As DayLightTime
    Set NewDayLightTime = New DayLightTime
    NewDayLightTime.Init StartTime, EndTime, Delta
End Function

''
' Returns a new <a href="OperatingSystem.html">OperatingSystem</a> object with the specified <i>Platform</i> and <i>Version</i>.
'
' @param Platform The platform the OS is.
' @param Version The version of the specified platform.
' @param ServicePack A message describing a service pack.
' @return The new <a href="OperatingSystem.html">OperatingSystem</a> object.
Public Function NewOperatingSystem(ByVal Platform As PlatformID, ByVal Version As Version, Optional ByRef ServicePack As String) As OperatingSystem
    Set NewOperatingSystem = New OperatingSystem
    NewOperatingSystem.Init Platform, Version, ServicePack
End Function

''
' Returns a new <a href="PlatformNotSupportedException.html">PlatformNotSupportedException</a> with the specified <i>Message</i> and <i>InnerException</i>.
'
' @param Message A human-readable error message.
' @param InnerException The exception that caused this exception.
' @return The new <a href="PlatformNotSupportedException.html">PlatformNotSupportedException</a> object.
Public Function NewPlatformNotSupportedException(ByRef Message As String, Optional ByVal InnerException As Exception) As PlatformNotSupportedException
    Set NewPlatformNotSupportedException = New PlatformNotSupportedException
    NewPlatformNotSupportedException.Init Message, InnerException
End Function

''
' Returns a new <a href="PathTooLongException.html">PathTooLongException</a> object.
'
' @param Message A human-readable message.
' @param InnerException The exception that caused this exception.
' @return The new <a href="PathTooLongException.html">PathTooLongException</a>.
Public Function NewPathTooLongException(ByRef Message As String, Optional ByVal InnerException As Exception) As PathTooLongException
    Set NewPathTooLongException = New PathTooLongException
    NewPathTooLongException.Init Message, InnerException
End Function

''
' Returns a new <a href="UTF8Encoding.html">UTF8Encoding</a> object.
'
' @param ShouldEmitUTF8Identifier <b>True</b> to specify that a Unicode byte order mark is provided; otherwise, <b>False</b>.
' @param ThrowOnInvalidBytes <b>True</b> to specify that an exception be thrown when an invalid encoding is detected; otherwise, <b>False</b>.
' @return The new <a href="UTF8Encoding.html">UTF8Encoding</a> object.
' @remarks If <i>ThrowOnInvalidBytes</i> is <b>True</b>, a method that detects an invalid byte sequence throws <a href="ArgumentException.html">ArgumentException</a>.
' Otherwise, the method does not throw an exception, and the invalid sequence is ignored.
Public Function NewUTF8Encoding(ByVal ShouldEmitUTF8Identifier As Boolean, Optional ByVal ThrowOnInvalidBytes As Boolean) As UTF8Encoding
    Set NewUTF8Encoding = New UTF8Encoding
    NewUTF8Encoding.Init ShouldEmitUTF8Identifier, ThrowOnInvalidBytes
End Function

''
' Returns a new <a href="UTF7Encoding.html">UTF7Encoding</a> object.
'
' @param AllowOptionals <b>True</b> to specify that optional characters are allowed; otherwise, <b>False</b>.
' @return The new <a href="UTF7Encoding.html">UTF7Encoding</a> object.
' @remarks <p>If an instance allows optional characters, Unicode code points are encoded with a corresponding optional
' character instead of a modified base 64 character. The optional characters are exclamation point ("!"), backward slash ("\"),
' vertical line ("|"), double quote ("""), number sign ("#"), dollar sign ("$"), percent sign ("%"), ampersand ("&"), asterisk ("*"),
' semicolon (";"), left angle bracket ("&lt;"), right angle bracket ("&gt;"), left curly bracket ("{"), right curly bracket ("}"),
' left square bracket ("["), right square bracket ("]"), equal sign ("="), at sign ("@"), circumflex accent ("^"), underscore ("_"), and grave accent ("`").</p>
Public Function NewUTF7Encoding(ByVal AllowOptionals As Boolean) As UTF7Encoding
    Set NewUTF7Encoding = New UTF7Encoding
    NewUTF7Encoding.Init AllowOptionals
End Function

''
' Returns a new <a href="ObjectDisposedException.html">ObjectDisposedException</a>.
'
' @param ObjectName The name of the object that has been disposed of.
' @param Message A human-readable message.
' @return The newly created <a href="ObjectDisposedException.html">ObjectDisposedException</a> object.
Public Function NewObjectDisposedException(ByRef ObjectName As String, Optional ByRef Message As String) As ObjectDisposedException
    Set NewObjectDisposedException = New ObjectDisposedException
    NewObjectDisposedException.Init ObjectName, Message
End Function

''
' Initializes a new instance of the <a href="MemoryStream.html">MemoryStream</a> class based on the specified region of a byte array, with the CanWrite property set as specified.
'
' @param Buffer The array of bytes from which to create this stream.
' @param Index The index in <i>Buffer</i> at which the stream begins. The default is lower bound of <i>Buffer</i>.
' @param Count The length of the stream in bytes. The default is <i>Buffer</i> length.
' @param Writable The setting of the <a href="MemoryStream.Get.CanWrite.html">CanWrite</a> property, which determines whether the stream supports writing.
' @param Mode A <a href="MemoryStream.BufferMode.html">BufferMode</a> value specifying how <i>Buffer</i> is handled internally.
' @return The newly created <a href="MemoryStream.html">MemoryStream</a> containing the supplied byte buffer.
' @remarks The <a href="MemoryStream.Get.CanRead.html">CanRead</a> and <a href="MemoryStream.GetCanSeek.html">CanSeek</a> properties
' are both set to <b>True</b>. <a href="MemoryStream.Let.Capacity.html">Capacity</a> is set to <i>Count</i>.
' <p>The length of the stream cannot be set to a value greater than the initial length of the specified byte array; however, the stream can be truncated (see <a href="MemoryStream.SetLength.html">SetLength</a>).</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Buffer</i> is uninitialized.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>Index</i> or <i>Count</i> are negative.</td>
'    </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The buffer length minus <i>Index</i> is less than <i>Count</i>.</td>
'    </tr>
' </table>
Public Function NewMemoryStream(ByRef Buffer() As Byte, Optional ByRef Index As Variant, Optional ByRef Count As Variant, Optional ByVal Writable As Boolean = True, Optional ByVal Mode As BufferMode = ShareMode) As MemoryStream
    Set NewMemoryStream = New MemoryStream
    NewMemoryStream.Init Buffer, Index, Count, Writable, Mode
End Function

''
' Returns a new <a href="StringBuilder.html">StringBuilder</a> with user specified settings and starting string.
'
' @param s A string to be inititially placed in the builder. If this is specified, then <a href="StringBuilder.set.Capacity.html">Capacity</a> will be set to no less than <i>Count</i>.
' @param StartIndex The starting index in <i>s</i> to begin placing into the buffer.
' @param Count The number of characters in <i>s</i> to place into the buffer.
' @param Capacity The capacity of the internal buffer. If this is less than <i>Count</i>, then <i>Count</i> will override this value.
' @return A new <a href="StringBuilder.html">StringBuilder</a> with <i>s</i> already placed in the buffer.
Public Function NewStringBuilder(Optional ByRef s As String, Optional ByRef StartIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Capacity As Long = 128) As StringBuilder
    Set NewStringBuilder = New StringBuilder
    NewStringBuilder.Init s, StartIndex, Count, Capacity
End Function

''
' Returns a new <a href="FileStream.html">FileStream</a> to access an underyling file.
'
' @param Path The name of the file to manipulate.
' @param Mode The method used to create or open a file.
' @param Access The access allowed to the file for this <a href="FileStream.html">FileStream</a>.
' @param Share The access allowed to the file for other processes.
' @param BufferSize The number of bytes used to cache reads and writes from the physical file.
' @param UseAsync Flag used to open the file in asyncronous mode (currently not supported).
' @return The new <a href="FileStream.html">FileStream</a> opened to the specified file.
Public Function NewFileStream(ByRef Path As String, ByVal Mode As FileMode, Optional ByVal Access As FileAccess, Optional ByVal Share As FileShare = FileShare.ReadShare, Optional ByVal BufferSize As Long = 4096, Optional ByVal UseAsync As Boolean = False) As FileStream
    If Access = 0 Then
        If Mode = FileMode.Append Then
            Access = FileAccess.WriteAccess
        Else
            Access = FileAccess.ReadWriteAccess
        End If
    End If
    
    Set NewFileStream = New FileStream
    NewFileStream.Init Path, Mode, Access, Share, BufferSize, UseAsync
End Function

''
' Returns a new <a href="FileStream.html">FileStream</a> to access a file other than that on disk, such as pipes.
'
' @param Handle A <a href="SafeFileHandle.html">SafeFileHandle</a> to the already opened filed.
' @param Access The access allowed to the file for this <a href="FileStream.html">FileStream</a>.
' @param BufferSize The number of bytes used to cache reads and writes from the physical file.
' @return The new <a href="FileStream.html">FileStream</a> opened to the specified file.
Public Function NewFileStreamWithHandle(ByVal Handle As SafeFileHandle, ByVal Access As FileAccess, Optional ByVal BufferSize As Long = 4096) As FileStream
    Set NewFileStreamWithHandle = New FileStream
    NewFileStreamWithHandle.InitWithHandle Handle, Access, BufferSize
End Function

''
' Returns a new <a href="UnicodeEncoding.html">UnicodeEncoding</a> object with the specified settings.
'
' @param BigEndian <b>True</b> to use the big endian byte order (most significant byte first), or <b>False</b> to use the little endian byte order (least significant byte first).
' @param ByteOrderMark <b>True</b> to specify that the <a href="UnicodeEncoding.GetPreamble.html">GetPreamble</a> method returns a Unicode byte order mark; otherwise, <b>False</b>. See the Remarks section for more information.
' @param ThrowOnInvalidBytes <b>True</b> to specify that an exception should be thrown when an invalid encoding is detected; otherwise, <b>False</b>.
' @return The new UnicodeEncoding object.
' @remarks If <i>ByteOrderMark</i> is <b>True</b> then the <a href="UnicodeEncoding.GetPreamble.html">GetPreamble</a> will emit one of the following byte arrays.
' <p>BigEnding = True: &HFE, &HFF<br>
' BigEnding = False: &HFF, &HFE</p>
' <p>If <i>ByteOrderMark</i> is <b>False</b> then <a href="UnicodeEncoding.GetPreamble.html">GetPreamble</a> will emit an empty array.</p>
Public Function NewUnicodeEncoding(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean, Optional ByVal ThrowOnInvalidBytes As Boolean) As UnicodeEncoding
    Set NewUnicodeEncoding = New UnicodeEncoding
    NewUnicodeEncoding.Init BigEndian, ByteOrderMark, ThrowOnInvalidBytes
End Function

''
' Initializes a new instance of the <a href="UTF32Encoding.html">UTF32Encoding</a> class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected.
'
' @param BigEndian <b>True</b> to use the big endian byte order (most significant byte first), or <b>False</b> to use the little endian byte order (least significant byte first).
' @param ByteOrderMark <b>True</b> to specify that a Unicode byte order mark is provided; otherwise, <b>False</b>.
' @param ThrowOnInvalidBytes <b>True</b> to specify that an exception should be thrown when an invalid encoding is detected; otherwise, <b>False</b>.
' @return An initialized <a href="UTF32Encoding.html">UTF32Encoding</a> object.
Public Function NewUTF32Encoding(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean, Optional ByVal ThrowOnInvalidBytes As Boolean) As UTF32Encoding
    Set NewUTF32Encoding = New UTF32Encoding
    NewUTF32Encoding.Init BigEndian, ByteOrderMark, ThrowOnInvalidBytes
End Function

''
' Returns a new <a href="StringWriter.html">StringWriter</a> object.
'
' @param sb A supplied <a href="StringBuilder.html">StringBuilder</a> object that is written to. This allows the <a href="StringWriter.html">StringWriter</a> to start with an already created string.
' @param Provider An object that provides formatting information for numbers and dates.
' @return The new <a href="StringWriter.html">StringWriter</a> object.
Public Function NewStringWriter(Optional ByVal sb As StringBuilder, Optional ByVal Provider As IFormatProvider) As StringWriter
    Set NewStringWriter = New StringWriter
    NewStringWriter.Init sb, Provider
End Function

''
' Returns a new <a href="StringReader.html">StringReader</a> object.
'
' @param s The string to be read from.
' @param CopyString Indicates that <a href="StringReader.html">StringReader</a> should copy the string or reference the original.
' @return The new <a href="StringReader.html">StringReader</a> object.
' @remarks By default the parameter <i>s</i> is referenced internally. This requires the original holder of the
' string variable to remain valid until <a href="StringReader.html">StringReader</a> has been disposed. If the
' original string cannot be maintained, then passing <b>True</b> for <i>CopyString</i> will cause <a href="StringReader.html">StringReader</a>
' to make a copy of the source string to prevent any string deallocation issues.
Public Function NewStringReader(ByRef s As String, Optional ByVal CopyString As Boolean) As StringReader
    Set NewStringReader = New StringReader
    NewStringReader.Init s, CopyString
End Function

''
' Creates a new <a href="DirectoryInfo.html">DirectoryInfo</a> object that describes the specified path.
'
' @param Path The directory path this instance describes and manipulates.
' @return The new object used to manage a directory path.
' @remarks When a <a href="DirectoryInfo.html">DirectoryInfo</a> object is created, there is no
' attempt at verifying the <i>Path</i>. Only when an attempt to manipulate the actual directory is the path validated.
Public Function NewDirectoryInfo(ByRef Path As String) As DirectoryInfo
    Set NewDirectoryInfo = New DirectoryInfo
    NewDirectoryInfo.Init Path
End Function

''
' Creates a new <a href="DirectoryNotFoundException.html">DirectoryNotFoundException</a> object with the specified message and optional exception.
'
' @param Message A human-readable message describing the exception.
' @param InnerException The exception that caused this exception to be thrown.
' @return A new <a href="DirectoryNotFoundException.html">DirectoryNotFoundException</a> object with the specified message and optional exception.
Public Function NewDirectoryNotFoundException(ByRef Message As String, Optional ByVal InnerException As Exception) As DirectoryNotFoundException
    Set NewDirectoryNotFoundException = New DirectoryNotFoundException
    NewDirectoryNotFoundException.Init Message, InnerException
End Function

''
' Creates a new <a href="StreamWriter.html">StreamWriter</a> with either a filename or other stream as a source to write to in the specified <a href="Encoding.html">Encoding</a>.
'
' @param Source Either a file name or <a href="Stream.html">Stream</a> object to write to.
' @param Append Indicates if a file is to be appended to or overwritten.
' @param Encoding The encoding to be used when writing to the stream.
' @param BufferSize The minimum size of the internal buffer used to cache writes.
' @param LeaveOpen <b>True</b> to leave the stream open after the <a href="StreamWriter.html">StreamWriter</a> object is disposed; otherwise, <b>False</b>.
' @return A new <a href="StreamWriter.html">StreamWriter</a> ready to write to an underlying stream in the specified <a href="Encoding.html">Encoding</a>.
' @remarks <p>If a file name is passed in, then a <a href="FileStream.html">FileStream</a> is created internally using either
' <a href="FileStream.FileMode.html">FileMode.Create</a> or <a href="FileStream.FileMode.html">FileMode.Append</a> depending on <i>Append</i>.</p>
' <p>If <i>Encoding</i> is not supplied, then an encoding of <a href="Constructors.NewUTF8Encoding.html">NewUTF8Encoding(False, True)</a> is used.</p>
Public Function NewStreamWriter(ByRef Source As Variant, Optional ByVal Append As Boolean = False, Optional ByVal Encoding As Encoding, Optional ByVal BufferSize As Long = 4096, Optional ByVal LeaveOpen As Boolean) As StreamWriter
    Set NewStreamWriter = New StreamWriter
    NewStreamWriter.Init Source, Append, Encoding, BufferSize, LeaveOpen
End Function

''
' Creates a new <a href="StreamReader.html">StreamReader</a> from either a FileName or an existing stream to read from.
'
' @param Source A file name used to open an existing file as a <a href="FileStream.html">FileStream</a>, or an existing <a href="Stream.html">Stream</a> object.
' @param Encoding The encoding to be used when decoding bytes from the stream.
' @param DetermineEncodingFromByteOrderMarks Requests that the reader attempt to determine what type of encoding is being used in the stream by reading the first few bytes.
' @param BufferSize The size of the internal cache used to improve performance.
' @param LeaveOpen A boolean flag indicating if the underlying stream is to be left open once the <a href="StreamReader.html">StreamReader</a> instance has closed.
' @return A newly initialized <a href="StreamReader.html">StreamReader</a> object.
' @remarks <p>If a file name is specified, then an internal <a href="FileStream.html">FileStream</a> object is created using an
' initialization of <a href="Constructors.NewFileStream.html">Cor.NewFileStream(Source, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare)</a>.</p>
' <p>If <i>DetermineEncodingFromByteOrderMarks</i> is <b>True</b>, then up to 3 bytes are read from the stream
' upon the first attempt to read any data from the stream. The bytes are used to determine if a specific
' encoding has been used. There are 3 encoding signatures that are looked for.
' <table style="width: auto">
'   <tr>
'       <th>BOM</th><th>Encoding</th>
'   </tr>
'   <tr>
'        <td>&HFE, &HFF</td><td><a href="UnicodeEncoding.html">UnicodeEncoding</a> with little-endian byte ordering is used.</td>
'   </tr>
'    <tr>
'        <td>&HFF, &HFE</td><td><a href="UnicodeEncoding.html">UnicodeEncoding</a> with big-endian byte ordering is used.</td>
'    </tr>
'    <tr>
'        <td>&HEF, &HBB, &HBF</td><td><a href="UTF8Encoding.html">UTF8Encoding</a> is used.</td>
'    </tr>
' </table>
' If no encoding is identified, then <a href="EncodingStatic.Get.UTF8.html">Encoding.UTF8</a> is used.
Public Function NewStreamReader(ByRef Source As Variant, Optional ByVal Encoding As Encoding, Optional ByVal DetermineEncodingFromByteOrderMarks As Boolean = True, Optional ByVal BufferSize As Long = 1024, Optional ByVal LeaveOpen As Boolean) As StreamReader
    Set NewStreamReader = New StreamReader
    NewStreamReader.Init Source, Encoding, DetermineEncodingFromByteOrderMarks, BufferSize, LeaveOpen
End Function

''
' Creates a new <a href="FileInfo.html">FileInfo</a> object used to describe and manage a file.
'
' @param FileName The name of the file to manage (may need path information as well.)
' @return The <a href="FileInfo.html">FileInfo</a> object to manage a file.
' @remarks When a <a href="FileInfo.html">FileInfo</a> object is first created, it does not verify the existence
' of the actual file. Once access to the file through the object is attempted, the file
' will be verified at that time. Portions of the object that doesn't directly interact
' with the file will not verify the existence of the file.
Public Function NewFileInfo(ByRef FileName As String) As FileInfo
    Set NewFileInfo = New FileInfo
    NewFileInfo.Init FileName
End Function

''
' Returns a new <a href="BinaryWriter.html">BinaryWriter</a> used to write to a Stream object.
'
' @param Stream The supplied stream.
' @param Encoding The character encoding.
' @param LeaveOpen <b>True</b> to leave the stream open after the <a href="BinaryWriter.html">BinaryWriter</a> object is disposed; otherwise, <b>False</b>.
' @return The new <a href="BinaryWriter.html">BinaryWriter</a> object.
' @remarks If <i>Encoding</i> is not supplied, then an encoding of <a href="Constructors.NewUTF8Encoding.html">Cor.NewUTF8Encoding(False, True)</a> is used.
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The stream does not support writing, or the stream is already closed.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Stream</i> is <b>Nothing</b>.</td>
'    </tr>
' </table>
Public Function NewBinaryWriter(ByVal Stream As Stream, Optional ByVal Encoding As Encoding, Optional ByVal LeaveOpen As Boolean) As BinaryWriter
    Set NewBinaryWriter = New BinaryWriter
    NewBinaryWriter.Init Stream, Encoding, LeaveOpen
End Function

''
' Returns a new <a href="EndOfStreamException.html">EndOfStreamException</a> with the specified message and optional exception.
'
' @param Message A human-readable message that describes the cause of the exception.
' @param InnerException The exception that caused this exception.
' @return The new <a href="EndOfStreamException.html">EndOfStreamException</a> object.
' @remarks This exception can be thrown whenever an attempt to read passed the end of a stream
' is attempted. Such classes as <a href="BinaryReader.html">BinaryReader</a> may throw this exception if necessary.
Public Function NewEndOfStreamException(ByRef Message As String, Optional ByVal InnerException As Exception) As EndOfStreamException
    Set NewEndOfStreamException = New EndOfStreamException
    NewEndOfStreamException.Init Message, InnerException
End Function

''
' Returns a new <a href="BinaryReader.html">BinaryReader</a> that can read from the specified stream usind the specified encoding.
'
' @param InputStream The stream to read the data from.
' @param Encoding The encoding system to be used to decode the bytes from the stream.
' @param LeaveOpen The stream can optionally be left open after the reader has been closed.
' @return The new <a href="BinaryReader.html">BinaryReader</a> with the specified stream and encoding.
' @remarks If <i>Encoding</i> is not specified, then the reader is created using <a href="UTF8Encoding.html">UTF8Encoding</a>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>InputStream</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>The stream does not support reading, or is already closed.</td>
'    </tr>
' </table>
' @include "..\DocIncludes\Constructors.NewBinaryReader.html"
Public Function NewBinaryReader(ByVal InputStream As Stream, Optional ByVal Encoding As Encoding, Optional ByVal LeaveOpen As Boolean) As BinaryReader
    Set NewBinaryReader = New BinaryReader
    NewBinaryReader.Init InputStream, Encoding, LeaveOpen
End Function

''
' Returns a new <a href="OverflowException.html">OverflowException</a> with the specified message and exception.
'
' @param Message A human-readable message describing the exception.
' @param InnerException The exception that caused this exception.
' @return The new <a href="OverflowException.html">OverflowException</a> object.
' @return This exception is thrown when a value may not fit into an existing datatype.
Public Function NewOverflowException(ByRef Message As String, Optional ByVal InnerException As Exception) As OverflowException
    Set NewOverflowException = New OverflowException
    NewOverflowException.Init Message, InnerException
End Function

''
' Creates an enumerator used to iterate over the Unicode characters of a <b>String</b>.
'
' @param s The string to iterate each character of.
' @param CopyString Indicates if the enumerator should create a copy of the passed in string, or use the existing string in memory.
' @return A <a href="CharEnumerator.html">CharEnumerator</a> used to iterate over each character of a string.
' @remarks <p>By default <a href="CharEnumerator.html">CharEnumerator</a> makes a copy of the passed in string.
' The original string can be changed as normal.</p>
' <p>If <i>CopyString</i> is set to <b>False</b>, then the original string must not change or go out of scope for the
' duration of the <a href="CharEnumerator.html">CharEnumerator</a> object's life time.</p>
' <p>By telling <a href="CharEnumerator.html">CharEnumerator</a> to not make a copy of the original string,
' memory is not allocated for the string, which can be helpful for large strings.</p>
Public Function NewCharEnumerator(ByRef s As String, Optional ByVal CopyString As Boolean) As CharEnumerator
    Set NewCharEnumerator = New CharEnumerator
    NewCharEnumerator.Init s, CopyString
End Function

''
' Creates a new instance of the <a href="GregorianCalendar.html">GregorianCalendar</a> class.
'
' @param CalendarType The language type the calendar is set to.
' @return A new instance of a <a href="GregorianCalendar.html">GregorianCalendar</a> class.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td><td><i>CalendarType</i> is not a member of the <a href="GregorianCalendar.GregorianCalendarTypes.html">GregorianCalendarTypes</a> enumeration.</td>
'   </tr>
' </table>
Public Function NewGregorianCalendar(ByVal CalendarType As GregorianCalendarTypes) As GregorianCalendar
    Set NewGregorianCalendar = New GregorianCalendar
    NewGregorianCalendar.Init CalendarType
End Function

''
' Creates a new <a href="ApplicationException.html">ApplicationException</a> object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <a href="ApplicationException.html">ApplicationException</a> object.
' @see ApplicationException
Public Function NewApplicationException(ByRef Message As String, Optional ByVal InnerException As Exception) As ApplicationException
    Set NewApplicationException = New ApplicationException
    NewApplicationException.Init Message, InnerException
End Function

''
' Creates a new <a href="ResourceReader.html">ResourceReader</a> object.
'
' @param Source The file name or stream to read from.
' @return A <a href="ResourceReader.html">ResourceReader</a> used to enumerate resources.
' @remarks <p>The <a href="ResourceReader.html">ResourceReader</a> reads the resources from a
' source with the same format as a .RES resource file.</p>
' <p><i>Source</i> should be either a file name or a <a href="Stream.html">Stream</a>
' to a set of resources in a .RES resource file.</p>
' <p>To read resources from a DLL or EXE use <a href="Constructors.NewWinResourceReader.html">NewWinResourceReader</a>.</p>
Public Function NewResourceReader(ByRef Source As Variant) As ResourceReader
    Set NewResourceReader = New ResourceReader
    NewResourceReader.Init Source
End Function

''
' Creates a new <a href="ResourceWriter.html">ResourceWriter</a> object.
'
' @param Source A file name or stream to write to.
' @return A <a href="ResourceWriter.html">ResourceWriter</a> used to write to a resource file.
' @remarks <p>The <a href="ResourceWriter.html">ResourceWriter</a> writes resources to a target
' in the same format as a .RES resource file.</p>
' <p><i>Source</i> should be either a file name or a <a href="Stream.html">Stream</a>
' to a set of resources in a .RES resource file.</p>
Public Function NewResourceWriter(ByRef Source As Variant) As ResourceWriter
    Set NewResourceWriter = New ResourceWriter
    NewResourceWriter.Init Source
End Function

''
' Creates a new <a href="CorDateTime.html">CorDateTime</a> object from a VB Date.
'
' @param Value The date.
' @param Kind Specifies whether the <a href="CorDateTime.html">CorDateTime</a> is a UTC, Local, or unspecified time type.
' @return A <a href="CorDateTime.html">CorDateTime</a> object.
' @remarks <p>This is a short-hand method for calling <a href="CorDateTimeStatic.FromOADate.html">CorDateTime.FromOADate(Value)</a>.</p>
' <pre>
' Dim d As CorDateTime
' Set d = NewDate(#1/1/2001#)
' </pre>
Public Function NewDate(ByVal Value As Date, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As CorDateTime
    Set NewDate = Statics.CorDateTime.FromOADate(Value, Kind)
End Function

''
' Creates a new <a href="ResourceSet.html">ResourceSet</a> for a specific culture.
'
' @param CultureID The specific culture to retrieve resources for.
' @param Source A .RES filename, <a href="IResourceReader.html">IResourceReader</a>, or <a href="Stream.html">Stream</a> object.
' @return A new <a href="ResourceSet.html">ResourceSet</a> object.
Public Function NewResourceSet(ByRef CultureID As Variant, ByRef Source As Variant) As ResourceSet
    Set NewResourceSet = New ResourceSet
    NewResourceSet.Init CultureID, Source
End Function

''
' Creates a new <a href="ResourceManager.html">ResourceManager</a> for a specified .RES file.
'
' @param Source The file name of the .RES file.
' @return A new <a href="ResourceManager.html">ResourceManager</a> object.
' @remarks If a set of resource files is to be loaded into a resource manager
' then use the <a href="ResourceManagerStatic.CreateFileBasedResourceManager.html">ResourceManager.CreateFileBasedResourceManager</a> method.
'
' @see ResourceManager
' @see ResourceManagerStatic
Public Function NewResourceManager(ByRef Source As String) As ResourceManager
    Set NewResourceManager = New ResourceManager
    NewResourceManager.InitFromFile Source
End Function

''
' Initializes a new instance of the <a href="Hashtable.html">Hashtable</a> class by copying the elements from the specified dictionary
' to the new <a href="Hashtable.html">Hashtable</a> object. The new <a href="Hashtable.html">Hashtable</a> object has an initial capacity equal to the number of
' elements copied, and uses the specified load factor and <a href="IEqualityComparer.html">IEqualityComparer</a> object.
'
' @param d The <a href="IDictionary.html">IDictionary</a> object containing keys and values to be added to the <a href="Hashtable.html">Hashtable</a>.
' @param LoadFactor A number in the range from 0.1 through 1.0 that is multiplied by the default value which provides the best performance. The result is the maximum ratio of elements to buckets.
' @param EqualityComparer The <a href="IEqualityComparer.html">IEqualityComparer</a> object that defines the hash code provider and the comparer to use with the <a href="Hashtable.html">Hashtable</a> object.
' @return A new <a href="Hashtable.html">Hashtable</a> object with the specified parameters.
' @see IEqualityComparer
' @see Hashtable
Public Function NewHashtable(Optional ByVal d As IDictionary, Optional ByVal LoadFactor As Single = 1!, Optional ByVal EqualityComparer As IEqualityComparer) As Hashtable
    Set NewHashtable = New Hashtable
    NewHashtable.Init d, 0, LoadFactor, EqualityComparer
End Function

''
' Initializes a new, empty instance of the <a href="Hashtable.html">Hashtable</a> class using the specified initial capacity, load factor, and <a href="IEqualityComparer.html">IEqualityComparer</a> object.
'
' @param Capacity The approximate number of elements that the <a href="Hashtable.html">Hashtable</a> object can initially contain.
' @param LoadFactor A number in the range from 0.1 through 1.0 that is multiplied by the default value which provides the best performance. The result is the maximum ratio of elements to buckets.
' @param EqualityComparer The <a href="IEqualityComparer.html">IEqualityComparer</a> object that defines the hash code provider and the comparer to use with the <a href="Hashtable.html">Hashtable</a> object.
' @return A new <a href="Hashtable.html">Hashtable</a> object with the specified parameters.
' @see IEqualityComparer
' @see Hashtable
Public Function NewHashtable2(ByVal Capacity As Long, Optional ByVal LoadFactor As Single = 1!, Optional ByVal EqualityComparer As IEqualityComparer) As Hashtable
    Set NewHashtable2 = New Hashtable
    NewHashtable2.Init Nothing, Capacity, LoadFactor, EqualityComparer
End Function

''
' Creates a new <a href="WinResourceReader.html">WinResourceReader</a> object.
'
' @param FileName The path and name of the EXE or DLL from which to read the resources.
' @return A new <a href="WinResourceReader.html">WinResourceReader</a> object.
Public Function NewWinResourceReader(ByRef FileName As String) As WinResourceReader
    Set NewWinResourceReader = New WinResourceReader
    NewWinResourceReader.Init FileName
End Function

''
' Initializes a new <a href="DriveInfo.html">DriveInfo</a> object used to inspect a drive's information.
'
' @param Name A valid drive path or drive letter. This can be either uppercase or lowercase, 'a' to 'z'.
' @return A <a href="DriveInfo.html">DriveInfo</a> object representing the specified drive.
' @remarks <p>Use this class to obtain information on drives. The drive name must be either an uppercase or lowercase letter
' from 'a' to 'z'. You cannot use this method to obtain information on drive names that are null or use UNC (\\server\share) paths.</p>
Public Function NewDriveInfo(ByRef Name As String) As DriveInfo
    Set NewDriveInfo = New DriveInfo
    NewDriveInfo.Init Name
End Function

''
' Initializes a new <a href="ConsoleKeyInfo.html">ConsoleKeyInfo</a> object that represents the state of a key press.
'
' @param KeyChar The Unicode character that corresponds to the <i>Key</i> parameter.
' @param Key The console key that corresponds to the <i>KeyChar</i> parameter.
' @param Shift <b>True<b/> to indicate that a SHIFT key was pressed; otherwise, <b>False</b>.
' @param Alt <b>True<b/> to indicate that a ALT key was pressed; otherwise, <b>False</b>.
' @param Control <b>True<b/> to indicate that a CTRL key was pressed; otherwise, <b>False</b>.
' @return A new <a href="ConsoleKeyInfo.html">ConsoleKeyInfo</a> object.
Public Function NewConsoleKeyInfo(ByVal KeyChar As Integer, ByVal Key As ConsoleKey, ByVal Shift As Boolean, ByVal Alt As Boolean, ByVal Control As Boolean) As ConsoleKeyInfo
    Set NewConsoleKeyInfo = New ConsoleKeyInfo
    NewConsoleKeyInfo.Init KeyChar, Key, Shift, Alt, Control
End Function

''
' Initializes a new <a href="TimeSpan.html">TimeSpan</a> representing the time portion of a date.
'
' @param DateValue The date to retrieve the time portion of.
' @return A <a href="TimeSpan.html">TimeSpan</a> object representing the time portion.
Public Function NewTime(ByRef DateValue As Variant) As TimeSpan
    Set NewTime = Convert.ToDateTime(DateValue).TimeOfDay
End Function

''
' Initializes a new <a href="IniFile.html">IniFile</a> to manage an INI file.
'
' @param FileName The INI filename.
' @return A new <a href="IniFile.html">IniFile</a> to manage an INI file.
Public Function NewIniFile(ByRef FileName As String) As IniFile
#If False Then
    Dim IniFile
#End If
    
    Set NewIniFile = New IniFile
    NewIniFile.Init FileName
End Function

''
' Creates a new <a href="Ticker.html">Ticker</a>.
'
' @param Interval The delay between events in milliseconds.
' @param Data A user supplied value to be associated with the timer.
' @param AutoReset <b>True</b> to have the timer event recur, <b>False</b> for one event.
' @param AddressOfCallback An address to a callback method that can be called when the timer has elapsed. Setting this to zero will stop the callback.
' @return A new <a href="Ticker.html">Ticker</a> object.
' @remarks <p>A callback method is used to allow a timer object to
' notify the function without having to use an Event. A function callback
' must have the following signature, or the application may crash.</p>
' <pre>
' Public Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
'     ' do stuff
' End Sub
' </pre>
' The two parameters must be declared as ByRef.
Public Function NewTicker(ByVal Interval As Long, Optional ByRef Data As Variant, Optional ByVal AutoReset As Boolean = True, Optional ByVal AddressOfCallback As Long) As Ticker
    Set NewTicker = New Ticker
    NewTicker.Init Interval, Data, AutoReset, AddressOfCallback
End Function

''
' Creates an array of values.
'
' @param ElementType The target datatype of the resultant array and the values to be converted to.
' @param ArrayElements The values to be put into the new array.
' @return An array of the specified datatype and containing the array elements with conversion,
' or an empty array if no elements are defined.
Public Function NewArray(ByVal ElementType As VbVarType, ParamArray ArrayElements() As Variant) As Variant
    Dim Length As Long
    Dim Args() As Variant
    
    Length = UBound(ArrayElements) + 1
    Helper.MoveVariant NewArray, CorArray.CreateInstance(ElementType, Length)
    
    If Length > 0 Then
        Helper.Swap4 ByVal ArrPtr(Args), ByVal Helper.DerefEBP(16)
        CorArray.CopyEx Args, 0, NewArray, 0, Length
    End If
End Function

''
' Creates a new <a href="ExceptionBase.html">ExceptionBase</a> object that can be contained within an exception class.
'
' @param Subclass The containing exception class that will contain this object instance.
' @param Message An understandable message that describes the exception being thrown.
' @param InnerException The exception that caused the current exception to be thrown.
' @param AppObject The App object of the application that contains the exception class.
' @return A new base class used for default handling of exception routines.
' @remarks Most exception classes simply contain this information to be retrieved for
' viewing. Few exception classes have additional requirements. This class can be used
' to handle the default set of information for an exception class. By containing this
' class inside the exception class, the exception class can delegate calls to this class
' for standard handling.
' <p>If additional values need to be maintained, then this class also supports the ability
' to set and get additional values using a string name as the key. The<a href="ArgumentException.html">ArgumentException</a>
' class uses this to hold the name of the parameter that caused the exception.</p>
' <p>This class is persistable and will also attempt to persist the custom values that
' have been added to it. If the persistence fails, then an error will occur.</p>
' @see ExceptionBase
Public Function NewExceptionBase(ByVal Subclass As Exception, Optional ByVal AppObject As Object) As ExceptionBase
    Set NewExceptionBase = New ExceptionBase
    NewExceptionBase.Init Subclass, AppObject
End Function

''
' Creates a new <a href="Queue.html">Queue</a> initialized with the elements in the collection.
'
' @param Source The elements to place in the queue upon creation.
' @return A <a href="Queue.html">Queue</a> object with the elements placed in it.
' @remarks The <i>Source</i> parameter can be a <b>VBA.Collection</b>, <a href="ICollection.html">ICollection</a> object, or an array.
Public Function NewQueue(ByRef Source As Variant) As Queue
    Set NewQueue = New Queue
    NewQueue.Init Source
End Function

''
' Creates a new <a href="ResourceKey.html">ResourceKey</a> object used to uniquely identify a resource.
'
' @param ResourceName The name of the resource being identified.
' @param ResourceType The type of resource being identified.
' @param LanguageID The Locale ID associated with the resource.
' @return An object that can be used to uniquely identify a resource.
' @remarks <i>ResourceName</i> and <i>ResourceType</i> can be a string or number.
Public Function NewResourceKey(ByRef ResourceName As Variant, ByRef ResourceType As Variant, ByVal LanguageID As Long) As ResourceKey
    Set NewResourceKey = New ResourceKey
    NewResourceKey.Init ResourceName, ResourceType, LanguageID
End Function

''
' Creates a picture group object used to contain information about resources that are grouped together.
'
' @param GroupType The type of resource the group represents.
' @return A group container for a specific resource type.
Public Function NewPictureResourceGroup(ByVal GroupType As PictureGroupTypes) As PictureResourceGroup
    Set NewPictureResourceGroup = New PictureResourceGroup
    NewPictureResourceGroup.Init GroupType
End Function

''
' Creates a new ExternalException object.
'
' @param Message A user-defined message specific to the exception.
' @param InnerException A reference to an exception that caused this exception.
' @return A new ExternalException object.
'
Public Function NewExternalException(ByRef Message As String, Optional ByVal InnerException As Exception) As ExternalException
    Set NewExternalException = New ExternalException
    NewExternalException.Init Message, InnerException
End Function

''
' Creates a new <a href="CspParameters.html">CspParameters</a> object.
'
' @param TypeIn The type code used to specify the kind of provider to create.
' @param ProviderNameIn The provider name.
' @param ContainerNameIn The provider container name.
' @return Returns a new <a href="CspParameters.html">CspParameters</a> object with the specified parameters.
Public Function NewCspParameters(ByVal TypeIn As Long, Optional ByRef ProviderNameIn As String, Optional ByRef ContainerNameIn As String) As CspParameters
    Set NewCspParameters = New CspParameters
    NewCspParameters.Init TypeIn, ProviderNameIn, ContainerNameIn
End Function

''
' Creates a new <a href="RNGCryptoServiceProvider.html">RNGCryptoServiceProvider</a> that uses a crypto service provider to generate random numbers.
'
' @param CspParams Parameters used to allow a specific crypto service provider to be used instead of the default.
' @return A new <a href="RNGCryptoServiceProvider.html">RNGCryptoServiceProvider</a> object that generated random numbers using a crypto service provider.
' @remarks Crypto-service providers generally produce a more random set of numbers than the standard
' Windows API set of random number generators. This is because cryptography depends a great deal on random number generation to be effective.
' <p>The default provider will be one of Microsoft's cryptography service providers such as
' "Microsoft Base Cryptographic Provider v1.0" or "Microsoft Enhanced Cryptographic Provider v1.0."</p>
' <p>Using the <i>CspParams</i> allows for a specific or custom provider to be used instead of the Microsoft providers.</p>
' @see RNGCryptoServiceProvider
Public Function NewRNGCryptoServiceProvider(ByVal CspParams As CspParameters) As RNGCryptoServiceProvider
    Set NewRNGCryptoServiceProvider = New RNGCryptoServiceProvider
    NewRNGCryptoServiceProvider.Init CspParams
End Function

''
' Creates a new <a href="CryptographicException.html">CryptographicException</a> object.
'
' @param Message The message contained within the exception object.
' @param InnerException A reference to an exception that may have cause this exception to occur.
' @return A new <a href="CryptographicException.html">CryptoGraphicException</a> object.
' @remarks Most of the exceptions that occur within the cryptography set of classes will be a <a href="CryptographicException.html">CryptographicException</a>.
' @see CryptographicException
Public Function NewCryptographicException(ByRef Message As String, Optional ByVal InnerException As Exception) As CryptographicException
    Set NewCryptographicException = New CryptographicException
    NewCryptographicException.Init Message, InnerException
End Function

''
' Creates a new <a href="KeySizes.html">KeySizes</a> object with the <a href="KeySizes.Get.MinSize.html">MinSize</a>,
' <a href="KeySizes.Get.MaxSize.html">MaxSize</a>, and <a href="KeySizes.Get.SkipSize.html">SkipSize</a> properties defined.
'
' @param MinSize The smallest key size defined by this <a href="KeySizes.html">KeySizes</a> instance.
' @param MaxSize The largest key size defined by this <a href="KeySizes.html">KeySizes</a> instance.
' @param SkipSize The step size from <a href="KeySizes.Get.MinSize.html">MinSize</a> to <a href="KeySizes.Get.MaxSize.html">MaxSize</a> of legal key sizes
' between the smallest and largest legal key sizes in this instance.
' @return A new <a href="KeySizes.html">KeySizes</a> object with the <a href="KeySizes.Get.MinSize.html">MinSize</a>, <a href="KeySizes.Get.MaxSize.html">MaxSize</a>, and <a href="KeySizes.Get.SkipSize.html">SkipSize</a> properties defined.
' @remarks A <a href="KeySizes.html">KeySizes</a> object represents a set of legal key sizes. It defines all the possible
' legal key sizes by defining the smallest and largest key sizes allows, then defining a step
' size used to calculate the additional legal key sizes between the smallest and largest key sizes.
' <p>If the smallest key size is 64 and the largest key size is 128 and a skip size of 32 is defined,
' then all possible legal key sizes are from 64 to 128 in increments of 32. In this case, the valid
' key sizes are 64, 96 and 128.</p>
' @see KeySizes
Public Function NewKeySizes(ByVal MinSize As Long, ByVal MaxSize As Long, ByVal SkipSize As Long) As KeySizes
    Set NewKeySizes = New KeySizes
    NewKeySizes.Init MinSize, MaxSize, SkipSize
End Function

''
' Creates a new <a href="FromBase64Transform.html">FromBase64Transform</a> object used to convert base64 characters to plain text.
'
' @param WhiteSpaces Determines how whitespace should be handled during the transformation process.
' @return A new <a href="FromBase64Transform.html">FromBase64Transfrom</a> object.
' @see FromBase64Transform
Public Function NewFromBase64Transform(ByVal WhiteSpaces As FromBase64TransformMode) As FromBase64Transform
    Set NewFromBase64Transform = New FromBase64Transform
    NewFromBase64Transform.Init WhiteSpaces
End Function

''
' Creates a new <a href="CryptoStream.html">CryptoStream</a> object used to read or write transformed data to or from an underlying stream.
'
' @param Stream The stream on which to perform the cryptographic transformation.
' @param Transform The cryptographic transformation that is to be performed on the stream.
' @param Mode One of the <a href="CryptoStream.CryptoStreamMode.html">CryptoStreamMode</a> values.
' @return Initializes a <a href="CryptoStream.html">CryptoStream</a> object.
' @see CryptoStream
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'       <td><a href="ArgumentNullException.html">ArgumentNullException</a></td>
'       <td>
'           <i>Stream</i> is <b>Nothing</b><br/>
'           -or-<br/>
'           <i>Transform</i> is <b>Nothing</b>
'       </td>
'   </tr>
'   <tr>
'       <td><a href="ArgumentException.html">ArgumentException</a></td>
'       <td>
'           <i>Mode</i> is <a href="CryptoStream.CryptoStreamModel.html">ReadMode</a> and <i>Stream</i> is not readable.<br/>
'           -or-
'           <i>Mode</i> is <a href="CryptoStream.CryptoStreamModel.html">WriteMode</a> and <i>Stream</i> is not writable.<br/>
'           -or-<br/>
'           <i>Mode</i> is invalid.
'       </td>
'   </tr>
' </table>
Public Function NewCryptoStream(ByVal Stream As Stream, ByVal Transform As ICryptoTransform, ByVal Mode As CryptoStreamMode) As CryptoStream
    Set NewCryptoStream = New CryptoStream
    NewCryptoStream.Init Stream, Transform, Mode
End Function

''
' Initializes a <b>Long</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Long</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewLongs(ParamArray Values() As Variant) As Long()
    Dim Result() As Long
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbLong, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewLongs = Result
End Function

''
' Initializes an <a href="Int64Static.Int64.html">Int64</a> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <a href="Int64Static.Int64.html">Int64</a> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewInt64s(ParamArray Values() As Variant) As Int64()
    Dim Result() As Int64
    Dim Info As IRecordInfo
    Dim i As Long
    
    Set Info = GetRecordInfo(Statics.Int64.Zero)
    SAPtr(Result) = SafeArrayCreateVectorEx(vbUserDefinedType, 0, UBound(Values) + 1, ObjPtr(Info))
    
    For i = 0 To UBound(Values)
        Result(i) = CInt64(Values(i))
    Next
    
    NewInt64s = Result
End Function

''
' Initializes a <b>Byte</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Byte</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewBytes(ParamArray Values() As Variant) As Byte()
    Dim Result() As Byte
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbByte, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewBytes = Result
End Function

''
' Initializes an <b>Integer</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to an <b>Integer</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewIntegers(ParamArray Values() As Variant) As Integer()
    Dim Result() As Integer
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbInteger, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewIntegers = Result
End Function

''
' Initializes a <a href="BigInteger.html">BigInteger</a> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <a href="BigInteger.html">BigInteger</a> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewBigIntegers(ParamArray Values() As Variant) As BigInteger()
    Dim Result() As BigInteger
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVectorEx(vbObject, 0, UBound(Values) + 1, VarPtr(mBigIntegerGuid))
    
    For i = 0 To UBound(Values)
        Set Result(i) = CBigInt(Values(i))
    Next
    
    NewBigIntegers = Result
End Function

''
' Initializes an <b>Integer</b> array representing Unicode characters with elements copied from <i>Values</i>.
'
' @param Values A series of characters or their Unicode code points.
' @return An <b>Integer</b> array containing a set of Unicode characters.
' @remarks This is a convenience method to easily create an <b>Integer</b> array representing Unicode characters.
' <p>If a string with more than one character is processed, only the first character will be recognized.
' To convert an entire string to a character array use <a href="CorString.ToCharArray.html">CorString.ToCharArray</a>.</p>
' @see #NewString
' @see CorString
' @include "..\DocIncludes\Constructors.NewChars.html"
Public Function NewChars(ParamArray Values() As Variant) As Integer()
    Dim Result() As Integer
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbInteger, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        If VarType(Values(i)) = vbString Then
            If Len(Values(i)) <> 1 Then _
                Error.Format Format_NeedSingleChar
                
            Result(i) = AscW(Values(i))
        Else
            Result(i) = Char.ConvertFromInt32(Values(i))
        End If
    Next
    
    NewChars = Result
End Function

''
' Initializes a <b>String</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>String</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewStrings(ParamArray Values() As Variant) As String()
    Dim Result() As String
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbString, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        If VarType(Values(i)) = vbString Then
            Result(i) = Values(i)
        Else
            Result(i) = Object.ToString(Values(i))
        End If
    Next
    
    NewStrings = Result
End Function

''
' Initializes a <b>Single</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Single</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewSingles(ParamArray Values() As Variant) As Single()
    Dim Result() As Single
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbSingle, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewSingles = Result
End Function

''
' Initializes a <b>Double</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Double</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewDoubles(ParamArray Values() As Variant) As Double()
    Dim Result() As Double
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbDouble, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewDoubles = Result
End Function

''
' Initializes a <b>Decimal</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Decimal</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewDecimals(ParamArray Values() As Variant) As Variant()
    Dim Result() As Variant
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbVariant, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        If IsInt64(Values(i)) Then
            Result(i) = Statics.Int64.ToDecimal(CInt64(Values(i)))
        Else
            Result(i) = CDec(Values(i))
        End If
    Next
    
    NewDecimals = Result
End Function

''
' Initializes a <b>Date</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to <b>Date</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewDates(ParamArray Values() As Variant) As Date()
    Dim Result() As Date
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbDate, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewDates = Result
End Function

''
' Initializes a <b>Currency</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to <b>Currency</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewCurrencies(ParamArray Values() As Variant) As Currency()
    Dim Result() As Currency
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbCurrency, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewCurrencies = Result
End Function

''
' Initializes a <b>Boolean</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a <b>Boolean</b> or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewBooleans(ParamArray Values() As Variant) As Boolean()
    Dim Result() As Boolean
    Dim i As Long
    
    SAPtr(Result) = SafeArrayCreateVector(vbBoolean, 0, UBound(Values) + 1)
    
    For i = 0 To UBound(Values)
        Result(i) = Values(i)
    Next
    
    NewBooleans = Result
End Function

''
' Initializes a <b>Variant</b> array with elements copied from <i>Values</i>.
'
' @param Values The values to create the array from.
' @remarks This is the same as using the <b>Array</b> function.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
Public Function NewVariants(ParamArray Values() As Variant) As Variant()
    NewVariants = Values
End Function

''
' Creates a new <a href="CspKeyContainerInfo.html">CspKeyContainerInfo</a> object.
'
' @param Parameters The parameters used to derive information from.
' @return A new <a href="CspKeyContainerInfo.html">CspKeyContainerInfo</a> object.
' @remarks This object is usually created by crypto service provider. Instantiating an
' object directly may result in property throwing an exception when accessed.
' @see CspKeyContainerInfo
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Parameters</i> is <b>Nothing</b>.</td>
'   </tr>
' </table>
Public Function NewCspKeyContainerInfo(ByVal Parameters As CspParameters) As CspKeyContainerInfo
    Set NewCspKeyContainerInfo = New CspKeyContainerInfo
    NewCspKeyContainerInfo.Init Parameters
End Function

''
' Initializes a new instance of the <a href="Rfc2898DeriveBytes.html">Rfc2898DeriveBytes</a> class using a password, a salt or salt size, and number of iterations to derive the key.
'
' @param Password The password used to derive the key. This can be a <b>String</b> or a <b>Byte</b> array.
' @param Salt The key salt used to derive the key. This can be a <b>Byte</b> array or a numeric value indicating the size of the salt to be generated.
' @param Iterations The number of iterations for the operation.
' @return A new instance of the <a href="Rfc2898DeriveBytes.html">Rfc2898DeriveBytes</a> class.
' @include "..\DocIncludes\Constructors.NewRfc2898DeriveBytes.html"
Public Function NewRfc2898DeriveBytes(ByRef Password As Variant, ByRef Salt As Variant, Optional ByVal Iterations As Long = 1000) As Rfc2898DeriveBytes
    Set NewRfc2898DeriveBytes = New Rfc2898DeriveBytes
    NewRfc2898DeriveBytes.Init Password, Salt, Iterations
End Function

''
' Creates a <a href="HashAlgorithmBase.html">HashAlgorithmBase</a> object used by hash algorithms.
'
' @param Callback Used to call back into the specifics of an algorithm implementation.
' @param KeepWeakReference Indicates if the base class should maintain a weak reference to the
' <i>Callback</i> object to prevent a circular reference from occuring if the <i>Callback</i> object is the same
' object as the one instantiating this instance.
' @param UseBuffer Indicates if the internal buffering mechanism in the base class should be used,
' or if all data should be passed to the <i>Callback</i> object to allow custom buffering.
' @param BlockSize The size of the byte block that is processed at a time by the specific algorithm
' using this base class. This is to help with the buffering.
' @return A new <a href="HashAlgorithmBase.html">HashAlgorithmBase</a> object.
' @remarks This base class can handle most hash algorithm implementation requirement except for the
' specifics of a hash algorithm. That is where the <i>Callback</i> is used.
' @see HashAlgorithmBase
Public Function NewHashAlgorithmBase(ByVal Callback As IHashAlgorithm, Optional ByVal KeepWeakReference As Boolean, Optional ByVal UseBuffer As Boolean, Optional ByVal BlockSize As Long) As HashAlgorithmBase
    Set NewHashAlgorithmBase = New HashAlgorithmBase
    NewHashAlgorithmBase.Init Callback, KeepWeakReference, UseBuffer, BlockSize
End Function

''
' Creates a new <a href="CryptoAPIHash.html">CryptoAPIHash</a> object.
'
' @param HashAlg The algorithm of the hash to be used.
' @param HashSize The size of the hash output.
' @return A new <a href="CryptoAPIHash.html">CryptoAPIHash</a> object.
' @remarks A <a href="CryptoAPIHash.html">CryptoAPIHash</a> object is an implementation of the
' <a href="IHashAlgorithm.html">IHashAlgorithm</a> callback that uses the CryptoAPI hash routines.
' @see CryptoAPIHash
Public Function NewCryptoAPIHash(ByVal HashAlg As Long, ByVal HashSize As Long) As CryptoAPIHash
    Set NewCryptoAPIHash = New CryptoAPIHash
    NewCryptoAPIHash.Init HashAlg, HashSize
End Function

''
' Creates a new <a href="HMACSHA1.html">HMACSHA1</a> object.
'
' @param Key The key used to sign the message.
' @param UseManagedSha1 Indicates if the <a href="SHA1Managed.html">SHA1Managed</a> or <a href="SHA1CryptoServiceProvider.html">SHA1CryptoServiceProvider</a> should be used.
' @return A new <a href="HMACSHA1.html">HMACSHA1</a> object.
' @see HMACSHA1
Public Function NewHMACSHA1(ByRef Key() As Byte, Optional ByVal UseManagedSha1 As Boolean = False) As HMACSHA1
    Set NewHMACSHA1 = New HMACSHA1
    NewHMACSHA1.Init Key, UseManagedSha1
End Function

''
' Creates a <a href="HMACBase.html">HMACBase</a> object that can be used as a base implementation of an <a href="HMAC.html">HMAC</a> class.
'
' @param HashAlgorithm An instance of an algorithm to be used for Hash Message Authentication Code.
' @param KeySize The size of the key to be used for signing the message.
' @param HashName The name of the hash being used.
' @param Key The key used to sign the message.
' @return A new <a href="HMACBase.html">HMACBase</a> object.
' @remarks This class implements many of the common routines used by <a href="HMAC.html">HMAC</a> classes.
Public Function NewHMACBase(ByVal HashAlgorithm As HashAlgorithm, ByVal KeySize As Long, Optional ByRef HashName As String, Optional ByRef Key As Variant) As HMACBase
    Set NewHMACBase = New HMACBase
    NewHMACBase.Init HashAlgorithm, KeySize, HashName, Key
End Function

''
' Creates a new <a href="HMACSHA256.html">HMACSHA256</a> object.
'
' @param Key The key used to sign the message.
' @return A new <a href="HMACSHA256.html">HMACSHA256</a> object.
' @see HMACSHA256
Public Function NewHMACSHA256(ByRef Key() As Byte) As HMACSHA256
    Set NewHMACSHA256 = New HMACSHA256
    NewHMACSHA256.Init Key
End Function

''
' Creates a new <a href="HMACSHA384.html">HMACSHA384</a> object.
'
' @param Key The key used to sign the message.
' @return A new <a href="HMACSHA384.html">HMACSHA384</a> object.
' @see HMACSHA384
Public Function NewHMACSHA384(ByRef Key() As Byte) As HMACSHA384
    Set NewHMACSHA384 = New HMACSHA384
    NewHMACSHA384.Init Key
End Function

''
' Creates a new <a href="HMACSHA512.html">HMACSHA512</a> object.
'
' @param Key The key used to sign the message.
' @return A new <a href="HMACSHA512.html">HMACSHA512</a> object.
' @see HMACSHA512
Public Function NewHMACSHA512(ByRef Key() As Byte) As HMACSHA512
    Set NewHMACSHA512 = New HMACSHA512
    NewHMACSHA512.Init Key
End Function

''
' Creates a new <a href="HMACMD5.html">HMACMD5</a> object.
'
' @param Key The key used to sign the message.
' @return A new <a href="HMACMD5.html">HMACMD5</a> object.
' @see HMACMD5
Public Function NewHMACMD5(ByRef Key() As Byte) As HMACMD5
    Set NewHMACMD5 = New HMACMD5
    NewHMACMD5.Init Key
End Function

''
' Creates a new <a href="HMACRIPEMD160.html">HMACRIPEMD160</a> object.
'
' @param Key The key used to sign the message.
' @return A new <a href="HMACRIPEMD160.html">HMACRIPEMD160</a> object.
' @see HMACRIPEMD160
Public Function NewHMACRIPEMD160(ByRef Key() As Byte) As HMACRIPEMD160
    Set NewHMACRIPEMD160 = New HMACRIPEMD160
    NewHMACRIPEMD160.Init Key
End Function

''
' Creates a new <a href="RSACryptoServiceProvider.html">RSACryptoServiceProvider</a> object.
'
' @param KeySize The size of the key used by the service provider.
' @param Parameters A <a href="CspParameters.html">CspParameters</a> object used to specify custom service provider options.
' @return A new <a href="RSACryptoServiceProvider.html">RSACryptoServiceProvider</a>.
' @remarks The key size can is 384bit to 512bit for non-high encryption machines and 384bit to 16384bit for high encryption machines.
' @see RSACryptoServiceProvider
Public Function NewRSACryptoServiceProvider(Optional ByVal KeySize As Long, Optional ByVal Parameters As CspParameters) As RSACryptoServiceProvider
    Set NewRSACryptoServiceProvider = New RSACryptoServiceProvider
    NewRSACryptoServiceProvider.Init KeySize, Parameters
End Function

''
' Creates a new <a href="DSACryptoServiceProvider.html">DSACryptoServiceProvider</a> object.
'
' @param KeySize The key size in bits. This can be from 384bits to 1024bits.
' @param Parameters A <a href="CspParameters.html">CspParameters</a> object used to specify custom service provider options.
' @return A new <a href="DSACryptoServiceProvider.html">DSACryptoServiceProvider</a> object.
' @see DSACryptoServiceProvider
Public Function NewDSACryptoServiceProvider(Optional ByVal KeySize As Long, Optional ByVal Parameters As CspParameters) As DSACryptoServiceProvider
    Set NewDSACryptoServiceProvider = New DSACryptoServiceProvider
    NewDSACryptoServiceProvider.Init KeySize, Parameters
End Function

''
' Creates a new <a href="SecurityElement.html">SecurityElement</a> object.
'
' @param Tag The name of the security element.
' @param Text The text within the element tags.
' @return A new <a href="SecurityElement.html">SecurityElement</a> object.
' @remarks A <a href="SecurityElement.html">SecurityElement</a> is a simple version of an XML tag with attributes and text.
' @see SecurityElement
Public Function NewSecurityElement(ByRef Tag As String, Optional ByRef Text As String) As SecurityElement
    Set NewSecurityElement = New SecurityElement
    NewSecurityElement.Init Tag, Text
End Function

''
' Creates a new <a href="XmlSyntaxException.html">XmlSyntaxException</a> object.
'
' @param LineNumber The number in the Xml string the exception occurred.
' @param Message The exception message.
' @param InnerException The exception that caused this exception to occur.
' @return A new <a href="XmlSyntaxException.html">XmlSyntaxException</a> object.
' @see XmlSyntaxException'
Public Function NewXmlSyntaxException(Optional ByRef LineNumber As Variant, Optional ByRef Message As Variant, Optional ByVal InnerException As Exception) As XmlSyntaxException
    Set NewXmlSyntaxException = New XmlSyntaxException
    NewXmlSyntaxException.Init LineNumber, Message, InnerException
End Function

''
' Creates a new <a href="MACTripleDES.html">MACTripleDES</a> object.
'
' @param RgbKey The key to be used by the implementation of the <a href="TripleDES.html">TripleDES</a> algorithm.
' @param StrTripleDES The name of the <a href="TripleDES.html">TripleDES</a> implementation to be used. The default is "3DES".
' @return A new <a href="MACTripleDES.html">MACTripleDES</a> object.
Public Function NewMACTripleDES(ByRef RgbKey() As Byte, Optional ByRef StrTripleDES As String) As MACTripleDES
    Set NewMACTripleDES = New MACTripleDES
    NewMACTripleDES.Init RgbKey, IfString(StrTripleDES, "3DES")
End Function

''
' Creates a new <a href="BigInteger.html">BigInteger</a> instance.
'
' @param value The value to intialize a new instance of <a href="BigInteger.html">BigInteger</a> to.
' @return A new <a href="BigInteger.html">BigInteger</a> instance containing the specified value.
' @remarks The follow datatypes are supported.
' <ul>
'   <li>vbLong</li>
'   <li>vbInteger</li>
'   <li>vbByte</li>
'   <li>vbDoube</li>
'   <li>vbSingle</li>
'   <li>vbCurrency</li>
'   <li>vbDecimal</li>
'   <li>vbByte Array</li>
'   <li><a href="Int64Static.Int64.html">Int64</a></li>
' </ul>
Public Function NewBigInteger(ByRef Value As Variant) As BigInteger
    Set NewBigInteger = New BigInteger
    NewBigInteger.Init Value
End Function

''
' Initializes a <b>Collection</b> with elements copied from <i>Values</i>.
'
' @param Values The values to initialize the collection with.
' @return Returns a collection containing items in <i>Values</i> parameter.
Public Function NewCollection(ParamArray Values() As Variant) As Collection
    Dim Value As Variant
    
    Set NewCollection = New Collection
    
    For Each Value In Values
        NewCollection.Add Value
    Next
End Function

''
' Initializes a new instance of the <a href="Guid.html">Guid</a> object by using the specified array of bytes, string, or memory pointer.
'
' @param Value An array of bytes, string, or memory pointer.
' @return A <a href="Guid.html">Guid</a> initialized with values from <i>Value</i>.
' @remarks <p>A new <a href="Guid.html">Guid</a> can be initialized from a specified array of bytes, string, or memory pointer.</p>
' <table style="width: auto">
'   <tr>
'       <th>Data Type</th><th>Description</th>
'   </tr>
'   <tr>
'        <td>String</td>
'        <td>
'           <p>A string that contains a GUID in one of the following formats ("d" represents a hexadecimal digit whose case is ignored):</p>
'           <p>32 contiguous Digits:</p>
'           <p>dddddddddddddddddddddddddddddddd</p>
'           <p>-or-</p>
'           <p>Groups of 8, 4, 4, 4, and 12 digits with hyphens between the groups. The entire GUID can optionally be enclosed in matching braces or parentheses:</p>
'           <p>dddddddd -dddd - dddd - dddd - dddddddddddd</p>
'           <p>-or-</p>
'           <p>{dddddddd-dddd-dddd-dddd-dddddddddddd}</p>
'           <p>-or-</p>
'           <p>(dddddddd-dddd-dddd-dddd-dddddddddddd)</p>
'           <p>-or-</p>
'           <p>Groups of 8, 4, and 4 digits, and a subset of eight groups of 2 digits, with each group prefixed by "0x" or "0X", and separated by commas. The entire GUID, as well as the subset, is enclosed in matching braces:</p>
'           <p>{0xdddddddd, 0xdddd, 0xdddd,{0xdd,0xdd,0xdd,0xdd,0xdd,0xdd,0xdd,0xdd}}</p>
'           <p>All braces, commas, and "0x" prefixes are required. All embedded spaces are ignored. All leading zeros in a group are ignored.</p>
'        </td>
'   </tr>
'    <tr>
'        <td>Byte Array</td><td>A 16-element byte array containing values with which to initialize the GUID.</td>
'    </tr>
'    <tr>
'        <td>Long</td><td>A pointer to a memory location to read 16 consecutive bytes with which to initialize the GUID.</td>
'    </tr>
' </table>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td>byte array is <b>null</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td>byte array is not 16 bytes long.</td>
'    </tr>
'    <tr>
'        <td><a href="FormatException.html">FormatException</a></td><td>string format is invalid.</td>
'    </tr>
' </table>
Public Function NewGuid(ByRef Value As Variant) As Guid
    Set NewGuid = New Guid
    NewGuid.Init Value
End Function

''
' Initializes a new instance of the <a href="SafeFileHandle.html">SafeFileHandle</a> class.
'
' @param PreexistingHandle An <b>Long</b> pointer that represents the pre-existing handle to use.
' @param OwnsHandle <b>True</b> to reliably release the handle during the finalization phase; <b>False</b> to prevent reliable release (not recommended).
' @return A new instance of the <a href="SafeFileHandle.html">SafeFileHandle</a> class.
Public Function NewSafeFileHandle(ByVal PreexistingHandle As Long, ByVal OwnsHandle As Boolean) As SafeFileHandle
    Set NewSafeFileHandle = New SafeFileHandle
    NewSafeFileHandle.Init PreexistingHandle, OwnsHandle
End Function

''
' Initializes a new instance of the <a href="SafeHandle.html">SafeHandle</a> class.
'
' @param PreexistingHandle An <b>Long</b> pointer that represents the pre-existing handle to use.
' @param OwnsHandle <b>True</b> to reliably release the handle during the finalization phase; <b>False</b> to prevent reliable release (not recommended).
' @param ZeroIsInvalid <b>True</b> if a handle of the value zero is considered an invalid handle; <b>False</b> otherwise.
' @return A new instance of the <a href="SafeHandle.html">SafeHandle</a> class.
' @remarks The <a href="SafeHandle.html">SafeHandle</a> class is used as base class through composition for managing specific handles.
Public Function NewSafeHandle(ByVal PreexistingHandle As Long, ByVal OwnsHandle As Boolean, Optional ByVal ZeroIsInvalid As Boolean = False) As SafeHandle
    Set NewSafeHandle = New SafeHandle
    NewSafeHandle.Init PreexistingHandle, OwnsHandle, ZeroIsInvalid
End Function

''
' Initializes a new instance of the <a href="DriveNotFoundException.html">DriveNotFoundException</a> class with the specified error message and a reference to the inner exception that is the cause of this exception.
'
' @param Message The error message that explains the reason for the exception.
' @param InnerException The exception that is the cause of the current exception. If the <i>InnerException</i> parameter is not <b>Nothing</b>, the current exception is raised in a <b>Catch</b> block that handles the inner exception.
Public Function NewDriveNotFoundException(ByRef Message As String, Optional ByVal InnerException As Exception) As DriveNotFoundException
    Set NewDriveNotFoundException = New DriveNotFoundException
    NewDriveNotFoundException.Init Message, InnerException
End Function

''
' Initializes a new instance of <a href="EnumeratorBase.html">EnumeratorBase</a>.
'
' @param StartIndex The starting position within a list of items to be enumerated.
' @param Length The number of items to be enumerated from the starting position.
' @param Collection An object that implements <a href="IVersionable.html">IVersionable</a> to track version changes of a collection.
' @return A base class used to manage the iteration of an index.
Public Function NewEnumeratorBase(ByVal StartIndex As Long, ByVal Length As Long, Optional ByVal Collection As IVersionable) As EnumeratorBase
    Set NewEnumeratorBase = New EnumeratorBase
    NewEnumeratorBase.Init StartIndex, Length, Collection
End Function

''
' Initializes a new instance of the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> class using a specified replacement string.
'
' @param Replacement A string that is emitted in a decoding operation in place of an input byte sequence that cannot be decoded.
' @return An initialized <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
' @remarks The <i>Replacement</i> parameter initializes the replacement string that is the value of the <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object. A commonly used value is the Unicode "Replacement Character" (U+FFFD), which is specifically intended to replace an incoming character having a value that is unknown or unrepresentable in Unicode.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Replacement</i> contains an invalid surrogate pair. In other words, the surrogate pair does not consist of one high surrogate component followed by one low surrogate component.</td>
'   </tr>
' </table>
Public Function NewDecoderReplacementFallback(ByRef Replacement As String) As DecoderReplacementFallback
    Set NewDecoderReplacementFallback = New DecoderReplacementFallback
    NewDecoderReplacementFallback.Init Replacement
End Function

''
' Initializes a new instance of the <a href="DecoderReplacementFallbackBuffer.html">DecoderReplacementFallbackBuffer</a> class using the value of a <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object.
'
' @param Fallback A <a href="DecoderReplacementFallback.html">DecoderReplacementFallback</a> object that contains a replacement string.
' @return An initialized <a href="DecoderReplacementFallbackBuffer.html">DecoderReplacementFallbackBuffer</a> object.
Public Function NewDecoderReplacementFallbackBuffer(ByVal Fallback As DecoderReplacementFallback) As DecoderReplacementFallbackBuffer
    Set NewDecoderReplacementFallbackBuffer = New DecoderReplacementFallbackBuffer
    NewDecoderReplacementFallbackBuffer.Init Fallback
End Function

''
' Initializes a new instance of the <a href="DecoderFallbackException.html">DecoderFallbackException</a> class. Parameters specify the error message and the inner exception that is the cause of this exception.
'
' @param Message An error message.
' @param InnerException The exception that caused this exception.
' @return A new <a href="DecoderFallbackException.html">DecoderFallbackException</a> instance.
Public Function NewDecoderFallbackException(ByRef Message As String, Optional ByVal InnerException As Exception) As DecoderFallbackException
    Set NewDecoderFallbackException = New DecoderFallbackException
    NewDecoderFallbackException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="DecoderFallbackException.html">DecoderFallbackException</a> class. Parameters specify the error message, the array of bytes being decoded, and the index of the byte that cannot be decoded.
'
' @param Message An error message.
' @param BytesUnknown The input byte array.
' @param Index The index position in <i>BytesUnknown</i> of the byte that cannot be decoded.
' @return A new <a href="DecoderFallbackException.html">DecoderFallbackException</a> instance.
Public Function NewDecoderFallbackException2(ByRef Message As String, ByRef BytesUnknown() As Byte, ByVal Index As Long) As DecoderFallbackException
    Set NewDecoderFallbackException2 = New DecoderFallbackException
    NewDecoderFallbackException2.Init2 Message, BytesUnknown, Index
End Function

''
' Initializes a new instance of the <a href="EncoderReplacementFallback.html">EncoderReplacementFallback</a> class using a specified replacement string.
'
' @param Replacement A string that is converted in an encoding operation in place of an input character that cannot be encoded.
' @remarks <p>The replacement parameter initializes the replacement string that is the value of this <a href="EncoderReplacementFallback.html">EncoderReplacementFallback</a> object.
' Your application must provide a replacement value that contains only characters that can be encoded in the target encoding.
' Otherwise, a recursive fallback results, causing an <a href="ArgumentException.html">ArgumentException</a>ArgumentException. For example, the fallback provided for an <a href="ASCIIEncoding.html">ASCIIEncoding</a> object
' cannot include the character "¿" (U+00BF) because that character is itself not a valid ASCII character.</p>
' <p>As a result of this, U+FFFD, which is a good choice for a fallback string for <a href="DecoderExceptionFallback.html">DecoderExceptionFallback</a>, is not generally a good choice for this class. Also, the null character (U+0000) cannot be used in the fallback string.</p>
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td><td><i>Replacement</i> contains an invalid surrogate pair. In other words, the surrogate does not consist of one high surrogate component followed by one low surrogate component.</td>
'   </tr>
' </table>
Public Function NewEncoderReplacementFallback(ByRef Replacement As String) As EncoderReplacementFallback
    Set NewEncoderReplacementFallback = New EncoderReplacementFallback
    NewEncoderReplacementFallback.Init Replacement
End Function

''
' Initializes a new instance of the <a href="EncoderReplacementFallbackBuffer.html">EncoderReplacementFallbackBuffer</a> class using the value of a EncoderReplacementFallback object.
'
' @param Fallback A <a href="EncoderReplacementFallback.html">EncoderReplacementFallback</a> object.
' @return A new instance of the <a href="EncoderReplacementFallbackBuffer.html">EncoderReplacementFallbackBuffer</a> class using the value of a EncoderReplacementFallback object.
Public Function NewEncoderReplacementFallbackBuffer(ByVal Fallback As EncoderReplacementFallback) As EncoderReplacementFallbackBuffer
    Set NewEncoderReplacementFallbackBuffer = New EncoderReplacementFallbackBuffer
    NewEncoderReplacementFallbackBuffer.Init Fallback
End Function

''
' Initializes a new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class. Parameters specify the error message and the inner exception that is the cause of this exception.
'
' @param Message An error message.
' @param InnerException The exception that caused this exception.
' @return A new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class.
Public Function NewEncoderFallbackException(ByRef Message As String, Optional ByVal InnerException As Exception) As EncoderFallbackException
    Set NewEncoderFallbackException = New EncoderFallbackException
    NewEncoderFallbackException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class. Parameters specify the error message, unknown character, and index.
'
' @param Message An error message.
' @param CharUnknown The unknown character that caused the exception.
' @param Index The index the unknown character was encountered.
' @return A new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class.
Public Function NewEncoderFallbackException2(ByRef Message As String, ByVal CharUnknown As Long, ByVal Index As Long) As EncoderFallbackException
    Set NewEncoderFallbackException2 = New EncoderFallbackException
    NewEncoderFallbackException2.Init2 Message, CharUnknown, Index
End Function

''
' Initializes a new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class. Parameters specify the error message, high surrogate, low surrogate, and Index.
'
' @param Message An error message.
' @param CharUnknownHigh The high surrogate of the unknown character.
' @param CharUnknownLow The low surrogate of the unknown character.
' @param Index The index the unknown character was encountered.
' @return A new instance of the <a href="EncoderFallbackException.html">EncoderFallbackException</a> class.
Public Function NewEncoderFallbackException3(ByRef Message As String, ByVal CharUnknownHigh As Long, ByVal CharUnknownLow As Long, ByVal Index As Long) As EncoderFallbackException
    Set NewEncoderFallbackException3 = New EncoderFallbackException
    NewEncoderFallbackException3.Init3 Message, CharUnknownHigh, CharUnknownLow, Index
End Function

''
' Initializes a new instance of the <a href="ArithmeticException.html">ArithmeticException</a> class with a specified error message and a reference to the inner exception that is the cause of this exception.
'
' @param Message The error message that explains the reason for the exception.
' @param InnerException The exception that is the cause of the current exception.
' @return A new instance of <a href="ArithmeticException.html">ArithmeticException</a>.
Public Function NewArithmeticException(ByRef Message As String, Optional ByVal InnerException As Exception) As ArithmeticException
    Set NewArithmeticException = New ArithmeticException
    NewArithmeticException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="DivideByZeroException.html">DivideByZeroException</a> class with a specified error message and a reference to the inner exception that is the cause of this exception.
'
' @param Message The error message that explains the reason for the exception.
' @param InnerException The exception that is the cause of the current exception.
' @return A new instance of <a href="DivideByZeroException.html">DivideByZeroException</a>.
Public Function NewDivideByZeroException(ByRef Message As String, Optional ByVal InnerException As Exception) As DivideByZeroException
    Set NewDivideByZeroException = New DivideByZeroException
    NewDivideByZeroException.Init Message, InnerException
End Function

''
' Initializes a new instance of the <a href="PictureResourceInfo.html">PictureResourceInfo</a> class with the specified <i>ResourceId</i> and <i>Resource</i>.
'
' @param ResourceId The resource number used to identify this resource.
' @param Resource The resource to be contained in the info.
' @return A new instance of the <a href="PictureResourceInfo.html">PictureResourceInfo</a>.
Public Function NewPictureResourceInfo(ByVal ResourceId As Long, ByVal Resource As IPicture) As PictureResourceInfo
    Set NewPictureResourceInfo = New PictureResourceInfo
    NewPictureResourceInfo.Init ResourceId, Resource
End Function

''
' Initializes a new instance of <a href="IniResourceWriter.html">IniResourceWriter</i> with specified <i>FileName</i> and <i>Section</i>.
'
' @param FileName The name of the ini file to be written to.
' @param Section The section within the ini file to be written to.
' @return A new instance of <a href="IniResourceWriter.html">IniResourceWriter</i>.
Public Function NewIniResourceWriter(ByRef FileName As String, ByRef Section As String) As IniResourceWriter
    Set NewIniResourceWriter = New IniResourceWriter
    NewIniResourceWriter.Init FileName, Section
End Function

''
' Returns an initialized instance of <a href="Int64Static.Int64.html">Int64</a>.
'
' @param LowPart The least significant 32-bits of the integer value.
' @param HighPart The most significant 32-bits of the integer value.
' @return An <a href="Int64Static.Int64.html">Int64</a> integer value.
' @remarks To convert a value to an <a href="Int64Static.Int64.html">Int64</a> use the <a href="PublicFunctions.CInt64.html">CInt64</a> conversion method.
Public Function NewInt64(ByVal LowPart As Long, ByVal HighPart As Long) As Int64
    NewInt64.HighPart = HighPart
    NewInt64.LowPart = LowPart
End Function

''
' Initializes a new <b>String</b> to the value indicated by an array of Unicode characters, a starting character position within that array, and a length.
'
' @param Value An array of Unicode characters.
' @param StartIndex The starting position within <i>Value</i>. Defaults to start of <i>Value</i>.
' @param Length The number of characters within <i>Value</i> to use. Defaults to using from <i>StartIndex</i> to end of array.
' @return A <b>String</b> initialized from the characters in <i>Value</i>.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Value</i> is <b>Nothing</b>.</td>
'   </tr>
'    <tr>
'        <td><a href="ArgumentOutOfRangeException.html">ArgumentOutOfRangeException</a></td>
'        <td>
'           <i>StartIndex</i> is less than the lower-bound of <i>Value</i>.<br/>
'           -or-<br/>
'           <i>Length</i> is less than zero.<br/>
'           -or-<br/>
'           The sum of <i>StartIndex</i> and <i>Length</i> is greater than the upper-bound of <i>Value</i>.
'        </td>
'    </tr>
' </table>
Public Function NewString(ByRef Value() As Integer, Optional ByRef StartIndex As Variant, Optional ByRef Length As Variant) As String
    Dim Range As ListRange
    
    ValidateArrayRange Value, StartIndex, Length, NameOfValue, NameOfStartIndex, NameOfLength
    Range = MakeArrayRange(Value, StartIndex, Length)
    
    If Range.Count = 0 Then
        NewString = ""
    Else
        NewString = SysAllocStringLen(VarPtr(Value(Range.Index)), Range.Count)
    End If
End Function

''
' Initializes a new instance of the <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> class using the specified <a href="CultureInfo.html">CultureInfo</a>.
'
' @param Culture The <a href="CultureInfo.html">CultureInfo</a> to use for the new <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a>.
' @return A new instance of the <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> class.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th><th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentNullException.html">ArgumentNullException</a></td><td><i>Culture</i> is <b>Nothing</b>.</td>
'   </tr>
' </table>
Public Function NewCaseInsensitiveComparer(ByVal Culture As CultureInfo) As CaseInsensitiveComparer
    Set NewCaseInsensitiveComparer = New CaseInsensitiveComparer
    NewCaseInsensitiveComparer.Init Culture
End Function

''
' Initializes a new instance of the <a href="Comparer.html">Comparer</a> class using the specified <a href="CultureInfo.html">CultureInfo</a>.
'
' @param Culture The <a href="CultureInfo.html">CultureInfo</a> to use for the new <a href="Comparer.html">Comparer</a>.
' @return A new instance of the <a href="Comparer.html">Comparer</a> class using the specified <a href="CultureInfo.html">CultureInfo</a>.
Public Function NewComparer(ByVal Culture As CultureInfo) As Comparer
    Set NewComparer = New Comparer
    NewComparer.Init Culture
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Internal Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function NewSafeMemoryMappedFileHandle(ByVal Handle As Long) As SafeMemoryMappedFileHandle
    Set NewSafeMemoryMappedFileHandle = New SafeMemoryMappedFileHandle
    NewSafeMemoryMappedFileHandle.Init Handle
End Function

Friend Function NewDefaultDecoder(ByVal Encoding As Encoding) As DefaultDecoder
    Set NewDefaultDecoder = New DefaultDecoder
    NewDefaultDecoder.Init Encoding
End Function

Friend Function NewDefaultEncoder(ByVal Encoding As Encoding) As DefaultEncoder
    Set NewDefaultEncoder = New DefaultEncoder
    NewDefaultEncoder.Init Encoding
End Function

Friend Function NewSBCSCodePageEncoding(ByVal CodePage As Long) As SBCSCodePageEncoding
    Set NewSBCSCodePageEncoding = New SBCSCodePageEncoding
    NewSBCSCodePageEncoding.Init CodePage
End Function

Friend Function NewDBCSCodePageEncoding(ByVal CodePage As Long) As DBCSCodePageEncoding
    Set NewDBCSCodePageEncoding = New DBCSCodePageEncoding
    NewDBCSCodePageEncoding.Init CodePage
End Function

Friend Function NewArrayListEnumerator(ByVal List As ArrayList, ByVal StartIndex As Long, ByVal Count As Long) As ArrayListEnumerator
    Set NewArrayListEnumerator = New ArrayListEnumerator
    NewArrayListEnumerator.Init List, StartIndex, Count
End Function

Friend Function NewMemoryMappedFile(ByVal Handle As SafeMemoryMappedFileHandle, Optional ByVal Stream As FileStream, Optional ByVal LeaveOpen As Boolean) As MemoryMappedFile
    Set NewMemoryMappedFile = New MemoryMappedFile
    NewMemoryMappedFile.Init Handle, Stream, LeaveOpen
End Function

Friend Function NewStreamAsyncResult(ByRef ObjectState As Variant) As StreamAsyncResult
    Set NewStreamAsyncResult = New StreamAsyncResult
    NewStreamAsyncResult.Init ObjectState
End Function

Friend Function NewRegistryKey(ByVal Key As Long, ByRef Name As String, ByVal Writable As Boolean) As RegistryKey
    Set NewRegistryKey = New RegistryKey
    NewRegistryKey.Init Key, Name, Writable
End Function

Friend Function NewSafeFindHandle(ByVal PreexistingHandle As Long, ByVal OwnsHandle As Boolean) As SafeFindHandle
    Set NewSafeFindHandle = New SafeFindHandle
    NewSafeFindHandle.Init PreexistingHandle, OwnsHandle
End Function

Friend Function NewOrdinalStringComparer(ByVal IgnoreCase As Boolean) As OrdinalStringComparer
    Set NewOrdinalStringComparer = New OrdinalStringComparer
    NewOrdinalStringComparer.Init IgnoreCase
End Function

Friend Function NewLegacyStringComparer(ByVal IgnoreCase As Boolean) As LegacyStringComparer
    Set NewLegacyStringComparer = New LegacyStringComparer
    NewLegacyStringComparer.Init IgnoreCase
End Function

Friend Function NewCultureAwareStringComparer(ByVal Culture As CultureInfo, ByVal IgnoreCase As Boolean) As CultureAwareStringComparer
    Set NewCultureAwareStringComparer = New CultureAwareStringComparer
    NewCultureAwareStringComparer.Init Culture, IgnoreCase
End Function

Friend Function NewCallbackSorter(ByVal Callback As Long) As CallbackSorter
    Set NewCallbackSorter = New CallbackSorter
    NewCallbackSorter.Init Callback
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    InitBigIntegerGuid
End Sub

Private Sub InitBigIntegerGuid()
' CopyMemory ByVal VarPtr(srcSA.SA) - IIDOffset, ByVal srcpSA - IIDOffset, LenB(srcSA.SA) + IIDOffset
    Dim Arr() As BigInteger
    
    ReDim Arr(0)
    CopyMemory mBigIntegerGuid, ByVal SAPtr(Arr) - vbSizeOfGuid, vbSizeOfGuid
End Sub

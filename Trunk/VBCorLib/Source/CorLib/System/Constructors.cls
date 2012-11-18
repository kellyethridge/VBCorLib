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
'    Module: Constructors
'

''
' Functions used to create fully inititalized objects.
'
' @remarks
' This class is defined as Global-Multiuse. All of the functions can be called
' directly without needing to create a <b>Constructors</b> object.
' <p>The <b>Cor</b> static object can be used to easily find and select the
' desired object to create. By using the <b>Cor.NewXxx</b> method, the
' IntelliSense will show all of the available constructors, making object
' creation easier.</p>
' <p><h4>Example</h4>
' <pre>
' '' This calls the constructor function directly.
' Set fs = NewFileStream("data.txt", FileMode.OpenOrCreate)
'
' '' This calls the constructor through the Cor.NewXxx method.
' Set fs = Cor.NewFileStream("data.txt", FileMode.OpenOrCreate)
' </pre>
' Both methods perform the same task, however, using the Cor.Newxxx method
' will provide IntelliSense showing all of the available constructors.
'
Option Explicit

''
' Creates a For..Each compatible wrapper around an IEnumerator object.
'
' @param obj The IEnumerator object that is to be called during For..Each.
' @return A For..Each compatible enumerator.
' @remarks <p><b>VBCorLib</b> allows for the building of custom enumerators to be used
' in a For..Each loop context.</p>
' <p>To create a custom enumerator please refer to the IEnumerator interface.</p>
' <p>In order for VB to utilize your custom IEnumerator object, it must be
' wrapped in an object that implements the IEnumVariant interface. This is the
' only interface that VB can use in a For..Each context. By passing in the
' custom enumerator into the <b>CreateEnumerator</b> function, a wrapper is applied and returned that
' is then returned in the typical NewEnum function of a class.
' <h4>Example</h4>
' <pre>
' Public Function NewEnum() As IUnknown
'     Dim en As New MyEnumerator
'     en.Init Me
'     Set NewEnum = CreateEnumerator(en)
' End Sub
' </pre>
' This example creates a custom enumerator object and initializes it with
' the object that is to be enumerated using the custom enumerator. The
' custom enumerator is then passed into the CreateEnumerator function and the
' returned value is then passed back to the calling function.
'
' @see IEnumerator
'
Public Function CreateEnumerator(ByVal Obj As IEnumerator) As IUnknown
    Set CreateEnumerator = modIEnumerator.CreateEnumerator(Obj)
End Function

''
' Creates a new Exception object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return The fully created Exception object.
' @see Exception
'
Public Function NewException(ByVal Message As String, Optional ByVal InnerException As Exception) As Exception
    Set NewException = New Exception
    Call NewException.Init(Message, InnerException)
End Function

''
' Creates a new SystemException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created SystemException object.
' @see SystemException
' @see Exception
'
Public Function NewSystemException(ByVal Message As String, Optional ByVal InnerException As Exception) As SystemException
    Set NewSystemException = New SystemException
    Call NewSystemException.Init(Message, InnerException)
End Function

''
' Creates a new UnauthorizedAccessException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created UnauthorizedAccessException object.
'
Public Function NewUnauthorizedAccessException(ByVal Message As String, Optional ByVal InnerException As Exception) As UnauthorizedAccessException
    Set NewUnauthorizedAccessException = New UnauthorizedAccessException
    Call NewUnauthorizedAccessException.Init(Message, InnerException)
End Function

''
' Creates a new NotSupportedException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created NotSupportedException object.
'
Public Function NewNotSupportedException(ByVal Message As String, Optional ByVal InnerException As Exception) As NotSupportedException
    Set NewNotSupportedException = New NotSupportedException
    Call NewNotSupportedException.Init(Message, InnerException)
End Function

''
' Creates a new ArgumentException object.
'
' @param Message A human-readable message explaining the exception.
' @param ParamName The parameter that caused the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created <b>ArgumentException</b> object.
' @remarks
' <p>When throwing an <b>ArgumentException</b> a meaningful error message describing the invalid argument, the name
' of the argument, as well as the expected range of values for the argument should be defined.</p>
' <h4>Example</h4>
' <p><pre>
' '' Throws a default ArgumentException object
' Throw New ArgumentException
'
' '' Throws an ArgumentException with a custom message and parameter name
' Throw Cor.NewArgumentException("Your argument is wrong.", "MyParam")
'
' '' Throws an ArgumentException with a custom message and error code and parameter name
' Dim ex As ArgumentException
' Set ex = Cor.NewArgumentException("Your argument is wrong.", "MyParam")
' ex.HResult = 123
' Throw ex
' </pre></p>
' This example shows how to throw three different variations of the same <b>ArgumentException</b> object.
'
' @see ArgumentException
'
Public Function NewArgumentException(ByVal Message As String, Optional ByVal ParamName As String, Optional ByVal InnerException As Exception) As ArgumentException
    Set NewArgumentException = New ArgumentException
    Call NewArgumentException.Init(Message, ParamName, InnerException)
End Function

''
' Creates a new ArgumentNullException object.
'
' @param Message A human-readable message explaining the exception.
' @param ParamName The parameter that caused the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created ArgumentNullException object.
' @remarks
' <p>When throwing an <b>ArgumentNullException</b> a meaningful error message describing the invalid argument
' and the name of the argument should be defined.</p>
' <h4>Example</h4>
' <p><pre>
' '' Throws a default ArgumentNullException object
' Throw New ArgumentNullException
'
' '' Throws an ArgumentNullException with a custom message and parameter name
' Throw Cor.NewArgumentNullException("A valid object is required.", "MyParam")
'
' '' Throws an ArgumentNullException with a custom message and error code and parameter name
' Dim ex As ArgumentNullException
' Set ex = Cor.NewArgumentNullException("A valid object is required.", "MyParam")
' ex.HResult = 123
' Throw ex
' </pre></p>
' This example shows how to throw three different variations of the same <b>ArgumentNullException</b> object.
'
' @see ArgumentNullException
'
Public Function NewArgumentNullException(ByVal Message As String, Optional ByVal ParamName As String, Optional ByVal InnerException As Exception) As ArgumentNullException
    Set NewArgumentNullException = New ArgumentNullException
    Call NewArgumentNullException.Init(Message, ParamName, InnerException)
End Function

''
' Creates a new ArgumentOutOfRangeException object.
'
' @param Message A human-readable message explaining the exception.
' @param ParamName The parameter that caused the exception.
' @param ActualValue The value of the parameter that caused the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created AgrumentOutOfRangeException object.
' @see ArgumentOutOfRangeException
'
Public Function NewArgumentOutOfRangeException(Optional ByVal Message As String, Optional ByVal ParamName As String, Optional ByVal ActualValue As Variant, Optional ByVal InnerException As Exception) As ArgumentOutOfRangeException
    Set NewArgumentOutOfRangeException = New ArgumentOutOfRangeException
    Call NewArgumentOutOfRangeException.Init(Message, ParamName, ActualValue, InnerException)
End Function

''
' Creates a new InvalidCaseException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created InvalidCastException object.
'
Public Function NewInvalidCastException(ByVal Message As String, Optional ByVal InnerException As Exception) As InvalidCastException
    Set NewInvalidCastException = New InvalidCastException
    Call NewInvalidCastException.Init(Message, InnerException)
End Function

''
' Creates a new IndexOutOfRangeException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created IndexOutOfRangeException object.
'
Public Function NewIndexOutOfRangeException(ByVal Message As String, Optional ByVal InnerException As Exception) As IndexOutOfRangeException
    Set NewIndexOutOfRangeException = New IndexOutOfRangeException
    Call NewIndexOutOfRangeException.Init(Message, InnerException)
End Function

''
' Creates a new RankException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created RankException object.
'
Public Function NewRankException(ByVal Message As String, Optional ByVal InnerException As Exception) As RankException
    Set NewRankException = New RankException
    Call NewRankException.Init(Message, InnerException)
End Function

''
' Creates a new InvalidOperationException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created InvalidOperationException object.
'
Public Function NewInvalidOperationException(ByVal Message As String, Optional ByVal InnerException As Exception) As InvalidOperationException
    Set NewInvalidOperationException = New InvalidOperationException
    Call NewInvalidOperationException.Init(Message, InnerException)
End Function

''
' Creates a new ArrayList object with the specified settings.
'
' @param comparer A comparer to be used for searching and sorting items in the list.
' @param c A collection to initialize the list with. Count will reflect the number of items added.
' @return A new ArrayList object.
' @remarks Sometimes the comparison of two items in the list cannot be accomplished
' by the default comparer that <b>ArrayList</b> uses. Cases like this may be when using
' different datatypes, UDT's, or objects that do not implement <b>IComparable</b>. In
' cases such as these, a custom comparer object may be supplied during the construction
' of an <b>ArrayList</b>. This comparer will be used as the default comparer for methods
' that require comparing items.
' <p>An <b>ArrayList</b> can be initialized with a default set of items. This is the same
' as creating a new <b>ArrayList</b> and using AddRange to add the items.</p>
' <p><b>ArrayList</b> accepts a VBA.Collection, ICollection, and an array as a list of
' items to initialize the new <b>ArrayList</b> to.
'
' <h4>Example</h4>
' <pre>
' Dim MyItems(10) As Long
'
' Set list = NewArrayList(New MyComparer, MyItems)
' </pre>
' This will create a new <b>ArrayList</b> that uses MyComparer as the default comparer
' for the items. The <b>ArrayList</b> is also initialized with an array of Longs. The
' fully initialized <b>ArrayList</b> will have a Count of 11 items, each item being
' a Long datatype with the value of 0.
'
' @see ArrayList
' @see IComparer
'
Public Function NewArrayList(Optional ByVal Comparer As IComparer, Optional ByRef c As Variant) As ArrayList
    Set NewArrayList = New ArrayList
    Call NewArrayList.Init(Comparer, c)
End Function

''
' Creates a new SerializationException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created SerializationException object.
'
Public Function NewSerializationException(ByVal Message As String, Optional ByVal InnerException As Exception) As SerializationException
    Set NewSerializationException = New SerializationException
    Call NewSerializationException.Init(Message, InnerException)
End Function

''
' Creates a new Stack object initialized with data in <i>col</i>.
'
' @param col A collection of items to initalize the Stack object with.
' @return A new Stack object containing the collection of values.
' @remarks <i>col</i> can be a VBA.Collection, an object implementing ICollection, or an Array.
'
Public Function NewStack(ByRef Col As Variant) As Stack
    Set NewStack = New Stack
    Call NewStack.Init(Col)
End Function

''
' Creates a new ArrayTypeMismatchException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created ArrayTypeMismatchException object.
'
Public Function NewArrayTypeMismatchException(ByVal Message As String, Optional ByVal InnerException As Exception) As ArrayTypeMismatchException
    Set NewArrayTypeMismatchException = New ArrayTypeMismatchException
    Call NewArrayTypeMismatchException.Init(Message, InnerException)
End Function

''
' Creates a new DictionaryEntry object initialized with <i>key</i> and <i>value</i>.
'
' @param Key A value designated as the key of the pair.
' @param value A value designated as the value of the pair.
' @return A new DictionaryEntry containing a key/value pair.
'
Public Function NewDictionaryEntry(ByRef Key As Variant, ByRef Value As Variant) As DictionaryEntry
    Set NewDictionaryEntry = New DictionaryEntry
    Call NewDictionaryEntry.Init(Key, Value)
End Function

''
' Creates a new SortedList object with specific initialization options.
'
' @param list An object implementing IDictionary used to initialize the list from.
' @param comparer An object implementing IComparer used to search and sort keys in the list.
' @param capacity An initial list size.
' @return A new SortedList object with the specific initializations.
'
Public Function NewSortedList(Optional ByVal list As IDictionary, Optional ByVal Comparer As IComparer, Optional ByVal Capacity As Long = 16) As SortedList
    Set NewSortedList = New SortedList
    Call NewSortedList.Init(list, Comparer, Capacity)
End Function

''
' Creates a new BitArray object.
'
' @param Length The number of bits the object will represent.
' @param value The value to initialize all the bits to (True = 1, False = 0).
' @return A new BitArray object set to the specified length with the specified value.
'
Public Function NewBitArray(ByVal Length As Long, Optional ByVal Value As Boolean = False) As BitArray
    Set NewBitArray = New BitArray
    Call NewBitArray.Init(Length, Value)
End Function

''
' Creates a new WeakReference object.
'
' @param Target The object to maintain a weak reference to.
' @return A new WeakReference object.
'
Public Function NewWeakReference(ByVal Target As Object) As WeakReference
    Set NewWeakReference = New WeakReference
    Call NewWeakReference.Init(Target)
End Function

''
' Creates a new FormatException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created FormatException object.
'
Public Function NewFormatException(ByVal Message As String, Optional ByVal InnerException As Exception) As FormatException
    Set NewFormatException = New FormatException
    Call NewFormatException.Init(Message, InnerException)
End Function

''
' Creates a new Random object initialized with the specified seed.
'
' @param seed The value used to initialize the random number generator with.
' @return A new Random object.
'
Public Function NewRandom(ByVal seed As Long) As Random
    Set NewRandom = New Random
    Call NewRandom.Init(seed)
End Function

''
' Creates a new Version object with specified version information.
'
' @param Major The number to describe an application level change.
' @param Minor The number that describes a component level change.
' @param Build The number used to describe an integration compiling of the application.
' @param Revision Minor adjustments to the application.
' @return A new Version object.
'
Public Function NewVersion(ByVal Major As Long, ByVal Minor As Long, Optional ByRef Build As Variant, Optional ByRef Revision As Variant) As Version
    Set NewVersion = New Version
    Call NewVersion.Init(Major, Minor, Build, Revision)
End Function

''
' Creates a new TimeSpan object set to the specified amount of time.
'
' @param Days The number of days in the timespan.
' @param Hours The number of hours in the timespan.
' @param Minutes The number of minutes in the timespan.
' @param Seconds The number of seconds in the timespan.
' @param Milliseconds The number of milliseconds in the timespan.
' @return A TimeSpan with the specified time.
' @remarks The values can be both positive and negative.
'
Public Function NewTimeSpan(Optional ByVal Days As Long, Optional ByVal Hours As Long, Optional ByVal Minutes As Long, Optional ByVal Seconds As Long, Optional ByVal Milliseconds As Long) As TimeSpan
    Set NewTimeSpan = New TimeSpan
    Call NewTimeSpan.Init(Days, Hours, Minutes, Seconds, Milliseconds)
End Function

''
' Creates a new cDateTime object set to the specified date.
'
' @param Year The year of the date.
' @param Month The month of the date.
' @param Day The day of the date.
' @param Hour The hour of the time portion of the date.
' @param Minute The minute of the time portion of the date.
' @param Second The second of the time portion of the date.
' @param Millisecond The millisecond of the time portion of the date.
' @param Calendar The calendar to use to create a cDateTime for a specific culture.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return A new cDateTime object initialized to the specified date and time.
'
Public Function NewcDateTime(Optional ByVal Year As Long, Optional ByVal Month As Long, Optional ByVal Day As Long, Optional ByVal Hour As Long, Optional ByVal Minute As Long, Optional ByVal Second As Long, Optional ByVal Millisecond As Long, Optional ByVal Calendar As Calendar, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set NewcDateTime = New cDateTime
    Call NewcDateTime.Init(Year, Month, Day, Hour, Minute, Second, Millisecond, Calendar, Kind)
End Function

''
' Creates a new IOException object with the specified Message, HResult, and InnerException.
'
' @param Message A human-readable error message.
' @param HResult The underlying error code.
' @param InnerException The exception that caused this exception.
' @return A new IOException object.
'
Public Function NewIOException(ByVal Message As String, Optional ByVal HResult As Long = COR_E_IO, Optional ByVal InnerException As Exception) As IOException
    Set NewIOException = New IOException
    Call NewIOException.Init(Message, HResult, InnerException)
End Function

''
' Creates a new FileNotFoundException with the specified Message, FileName, and InnerException.
'
' @param Message A human-readable error message.
' @param FileName The filename that caused the exception.
' @param InnerException The exception that caused this exception.
' @return A new FileNotFoundException with the specified values.
'
Public Function NewFileNotFoundException(Optional ByVal FileName As String, Optional ByVal Message As String, Optional ByVal InnerException As Exception) As FileNotFoundException
    Set NewFileNotFoundException = New FileNotFoundException
    Call NewFileNotFoundException.Init(FileName, Message, InnerException)
End Function

''
' Creates a memory-mapped view of a file.
'
' @param Path The file to be mapped to memory.
' @param Access The access to this mapping of the file.
' @param Share How should the file be shared with other processes.
' @return The new mapped view object.
'
Public Function NewMemoryMappedFile(ByVal Path As String, Optional ByVal Access As FileAccess = FileAccess.ReadAccess, Optional ByVal Share As FileShare = FileShare.ReadShare) As MemoryMappedFile
    Set NewMemoryMappedFile = New MemoryMappedFile
    Call NewMemoryMappedFile.Init(Path, Access, Share)
End Function

''
' Creates a new CultureInfo object for the specified culture.
'
' @param NameOrLCID The name or Locale ID of the culture to be created.
' @param UseUserOverride A flag indicating if the local machine settings should
' be used instead of the stored setting. This only applies to the culture of
' the current machine.
' @return The newly created CultureInfo object.
'
Public Function NewCultureInfo(ByRef NameOrLCID As Variant, Optional ByVal UseUserOverride As Boolean = True) As CultureInfo
    Set NewCultureInfo = New CultureInfo
    Call NewCultureInfo.Init(NameOrLCID, UseUserOverride)
End Function

''
' Creates a new DayLightTime object with the specified start and end times for daylight savings.
'
' @param StartTime The starting of the daylight savings period.
' @param EndTime The ending of the daylight savings period.
' @param Delta The time difference between daylight savings time and standard time.
' @return A newly created DayLightTime object.
'
Public Function NewDayLightTime(ByRef StartTime As Variant, ByRef EndTime As Variant, ByVal Delta As TimeSpan) As DayLightTime
    Set NewDayLightTime = New DayLightTime
    Call NewDayLightTime.Init(StartTime, EndTime, Delta)
End Function

''
' Returns a new OperatingSystem object with the specified Platform and Version.
'
' @param Platform The platform the OS is.
' @param Version The version of the specified platform.
' @param ServicePack A message describing a service pack.
' @return The new OperatingSystem object.
'
Public Function NewOperatingSystem(ByVal Platform As PlatformID, ByVal Version As Version, Optional ByVal ServicePack As String) As OperatingSystem
    Set NewOperatingSystem = New OperatingSystem
    Call NewOperatingSystem.Init(Platform, Version, ServicePack)
End Function

''
' Returns a new PlatformNotSupportedException with the specified Message and InnerException.
'
' @param Message A human-readable error message.
' @param InnerException The exception that caused this exception.
' @return The new PlatformNotSupportedException object.
'
Public Function NewPlatformNotSupportedException(ByVal Message As String, Optional ByVal InnerException As Exception) As PlatformNotSupportedException
    Set NewPlatformNotSupportedException = New PlatformNotSupportedException
    Call NewPlatformNotSupportedException.Init(Message, InnerException)
End Function

''
' Returns a new PathTooLongException object.
'
' @param Message A human-readable message.
' @param InnerException The exception that caused this exception.
' @return The new PathTooLongException.
'
Public Function NewPathTooLongException(ByVal Message As String, Optional ByVal InnerException As Exception) As PathTooLongException
    Set NewPathTooLongException = New PathTooLongException
    Call NewPathTooLongException.Init(Message, InnerException)
End Function

''
' Returns a new UTF8Encoding object.
'
' @param ShouldEmitUTF8Identifier Indicates whether or not the encoding should emit the preamble bytes in the GetPreamble method.
' @param ShouldThrowException Signals whether or not the encoding should throw an
' exception if the the source bytes are invalid for creating a character array.
' @param The new UTF8Encoding object.
'
Public Function NewUTF8Encoding(ByVal ShouldEmitUTF8Identifier As Boolean, Optional ByVal ShouldThrowException As Boolean) As UTF8Encoding
    Set NewUTF8Encoding = New UTF8Encoding
    Call NewUTF8Encoding.Init(ShouldEmitUTF8Identifier, ShouldThrowException)
End Function

''
' Returns a new UTF7Encoding object.
'
' @param AllowOptionals Signals whether or not to allow the optional character set of O from
' RFC2152 to be directly encoded or not.
' @return The new UTF7Encoding object.
'
Public Function NewUTF7Encoding(ByVal AllowOptionals As Boolean) As UTF7Encoding
    Set NewUTF7Encoding = New UTF7Encoding
    Call NewUTF7Encoding.Init(AllowOptionals)
End Function

''
' Returns a new ObjectDisposedException
'
' @param ObjectName The name of the object that has been disposed of.
' @param Message A human-readable message.
' @return The newly created ObjectDisposedException object.
'
Public Function NewObjectDisposedException(ByVal ObjectName As String, Optional ByVal Message As String) As ObjectDisposedException
    Set NewObjectDisposedException = New ObjectDisposedException
    NewObjectDisposedException.Init ObjectName, Message
End Function

' For internal use only.
Friend Function NewStreamAsyncResult(ByVal ObjectState As Variant) As StreamAsyncResult
    Set NewStreamAsyncResult = New StreamAsyncResult
    Call NewStreamAsyncResult.Init(ObjectState)
End Function

''
' Returns a new memoryStream object which uses a supplied byte array.
'
' @param Buffer The supplied byte array to act upon.
' @param Index The starting index in the array to begin usage.
' @param Count The total bytes that can be used by the stream.
' @param writable Whether or not to allow the stream to be written to.
' @return The newly created MemoryStream containing the supplied byte buffer.
' @remarks By using a user supplied buffer, the user can decide how much of the
' buffer is actually accessable. The buffer cannot be increased in size. The stream
' does not release the byte Buffer. The user still must maintain a reference to the
' array and keep it alive for the duration of the MemoryStream's life.
'
Public Function NewMemoryStream(ByRef Buffer() As Byte, Optional ByVal Index As Variant, Optional ByVal Count As Variant, Optional ByVal Writable As Boolean = True) As MemoryStream
    Set NewMemoryStream = New MemoryStream
    Call NewMemoryStream.Init(Buffer, Index, Count, Writable)
End Function

''
' Returns a new StringBuilder with user specified settings and starting string.
'
' @param s A string to be inititially placed in the builder. If this is specified, then
' Capacity will be set to no less than Count.
' @param StartIndex The starting index in s to begin placing into the buffer.
' @param Count The number of characters in s to place into the buffer.
' @param Capacity The capacity of the internal buffer. If this is less than Count, then
' Count will override this value.
' @return A new StringBuilder with s already placed in the buffer.
'
Public Function NewStringBuilder(Optional ByRef s As String, Optional ByRef startIndex As Variant, Optional ByRef Count As Variant, Optional ByVal Capacity As Long = 512) As StringBuilder
    Set NewStringBuilder = New StringBuilder
    Call NewStringBuilder.Init(s, startIndex, Count, Capacity)
End Function

''
' Returns a new FileStream to access an underyling file.
'
' @param Path The name of the file to manipulate.
' @param Mode The method used to create or open a file.
' @param Access The access allowed to the file for this FileStream.
' @param Share The access allowed to the file for other processes.
' @param BufferSize The number of bytes used to cache reads and writes from the physical file.
' @param UseAsync Flag used to open the file in asyncronous mode.
' @return The new FileStream opened to the specified file.
'
Public Function NewFileStream(ByVal Path As String, ByVal Mode As FileMode, Optional ByVal Access As FileAccess = -1, Optional ByVal Share As FileShare = FileShare.ReadShare, Optional ByVal BufferSize As Long = 4096, Optional ByVal UseAsync As Boolean = False) As FileStream
    Set NewFileStream = New FileStream
    Call NewFileStream.Init(Path, Mode, Access, Share, BufferSize, UseAsync)
End Function

''
' Returns a new FileStream to access a file other than that on disk, such as pipes.
'
' @param Handle The handle to the already opened filed.
' @param Access The access allowed to the file for this FileStream.
' @param OwnsHandle Determines if the FileStream object is responsible for closing the file passed in.
' @param BufferSize The number of bytes used to cache reads and writes from the physical file.
' @return The new FileStream opened to the specified file.
'
Public Function NewFileStreamFromHandle(ByVal Handle As Long, ByVal Access As FileAccess, Optional ByVal OwnsHandle As Boolean = True, Optional ByVal BufferSize As Long = 4096) As FileStream
    Set NewFileStreamFromHandle = New FileStream
    Call NewFileStreamFromHandle.InitFromHandle(Handle, Access, OwnsHandle, BufferSize)
End Function

''
' Returns a new UnicodeEncoding object with the specified settings.
'
' @param BigEndian Indicates if the byte order for a character should be high byte first (default is low byte first.)
' @param ByteOrderMark Indicates whether or not the encoding should emit the preamble bytes in the GetPreamble method.
' @return The new UnicodeEncoding object.
' @remarks If ByteOrderMark is True then the GetPreamble will emit one of the following byte arrays.
' <p>BigEnding = True: &HFE, &HFF<br>
' BigEnding = False: &HFF, &HFE</p>
'
Public Function NewUnicodeEncoding(ByVal BigEndian As Boolean, ByVal ByteOrderMark As Boolean) As UnicodeEncoding
    Set NewUnicodeEncoding = New UnicodeEncoding
    Call NewUnicodeEncoding.Init(BigEndian, ByteOrderMark)
End Function

''
' Returns a new StringWriter object.
'
' @param sb A supplied StringBuilder object that is written to. This allows the
' StringWriter to start with an already created string.
' @param provider An object that provided formatting information for numbers and dates.
' @return The new StringWriter object.
'
Public Function NewStringWriter(Optional ByVal sb As StringBuilder, Optional ByVal Provider As IFormatProvider) As StringWriter
    Set NewStringWriter = New StringWriter
    Call NewStringWriter.Init(sb, Provider)
End Function

''
' Returns a new StringReader object.
'
' @param s The string to be read from.
' @param IsOwner Tells StringReader to steal the string from the source.
' @return The new StringReader object.
' @remarks If <i>IsOwner</i> is True, then the variable passed into as <i>s</i>
' will become an empty string after the object is constructed.
'
Public Function NewStringReader(ByRef s As String, Optional ByVal IsOwner As Boolean) As StringReader
    Set NewStringReader = New StringReader
    Call NewStringReader.Init(s, IsOwner)
End Function

''
' Creates a new DirectoryInfo object that describes the specified path.
'
' @param Path The directory path this instance describes and manipulates.
' @return The new object used to manage a directory path.
' @remarks When a <b>DirectoryInfo</b> object is created, there is no
' attempt at verifying the <i>Path</i>. Only when an attempt to manipulate
' the actual directroy is the path validated.
'
Public Function NewDirectoryInfo(ByVal Path As String) As DirectoryInfo
    Set NewDirectoryInfo = New DirectoryInfo
    Call NewDirectoryInfo.Init(Path)
End Function

''
' Creates a new DirectoryNotFoundException object with the specified message and optional exception.
'
' @param Message A human-readable message describing the exception.
' @param InnerException The exception that caused this exception to be thrown.
' @return A new DirectoryNotFoundException object with the specified message and optional exception.
'
Public Function NewDirectoryNotFoundException(ByVal Message As String, Optional ByVal InnerException As Exception) As DirectoryNotFoundException
    Set NewDirectoryNotFoundException = New DirectoryNotFoundException
    Call NewDirectoryNotFoundException.Init(Message, InnerException)
End Function

''
' Creates a new StreamWriter with either a filename or other stream as a source to write to
' in the specified Encoding.
'
' @param Source Either a FileName or Stream object to write to.
' @param Encoding The encoding to be used when writing to the stream.
' @param BufferSize The minimum size of the internal buffer used to cache writes.
' @param Append Indicates if a file is to be appended to or overwritten.
' @return A new StreamWriter ready to write to an underlying stream in the specified Encoding.
' @remarks <p>If a FileName is passed in, then a FileStream is created internally Using either
' FileMode.Create or FileMode.Append depending on Append, FileAccess.WriteAccess, FileShare.ReadShare.</p>
' <p>If Encoding is not supplied, then an encoding of Cor.NewUTF8Encoding(False, True) is used.</p>
'
Public Function NewStreamWriter(ByVal Source As Variant, Optional ByVal Encoding As Encoding, Optional ByVal BufferSize As Long = 4096, Optional ByVal Append As Boolean = False) As StreamWriter
    Set NewStreamWriter = New StreamWriter
    Call NewStreamWriter.Init(Source, Encoding, BufferSize, Append)
End Function

''
' Creates a new StreamReader from either a FileName or an existing stream to read from.
'
' @param Source A FileName used to open an existing file as a FileStream, or an existing Stream object.
' @param Encoding The encoding to be used when decoding bytes from the stream.
' @param determineEncodingFromByteOrderMarks Requests that the reader attempt to determine what type of
' encoding is being used in the stream by reading the first few bytes.
' @param BufferSize The size of the internal cache used to improve performance.
' @return A newly initialized StreamReader object.
' @remarks <p>If a filename is specified, then an internal FileStream object is created using an
' initialization of Cor.NewFileStream(Source, FileMode.OpenExisting, FileAccess.ReadAccess, FileShare.ReadShare).</p>
' <p>If <i>determineEncodingFromByteOrderMarks</i> is True, then up to 3 bytes are read from the stream
' upon the first attempt to read any data from the stream. The bytes are used to determine if a specific
' encoding has been used. There are 3 encoding signatures that are looked for.<br>
' &HFE, &HFF - UnicodeEncoding with Little Endian byte ordering<br>
' &HFF, &HFE - UnicodeEncoding with Big Endian byte ordering<br>
' &HEF, &HBB, &HBF - UTF8Encoding<br>
' If no encoding is identified, then Encoding.UTF8 is used.
'
Public Function NewStreamReader(ByVal Source As Variant, Optional ByVal Encoding As Encoding, Optional ByVal determineEncodingFromByteOrderMarks As Boolean = True, Optional ByVal BufferSize As Long = 4096) As StreamReader
    Set NewStreamReader = New StreamReader
    Call NewStreamReader.Init(Source, Encoding, determineEncodingFromByteOrderMarks, BufferSize)
End Function

''
' Creates a new FileInfo object used to describe and manage a file.
'
' @param FileName The name of the file to manage (may need path information as well.)
' @return The FileInfo object to manage a file.
' @remarks When the FileInfo object is first created, it does not verify the existence.
' of the actual file. Once access to the file through the object is attempt, the file
' will be verified at that time. Portions of the object that doesn't directly interact
' with the file will not verify the existence of the file.
'
Public Function NewFileInfo(ByVal FileName As String) As FileInfo
    Set NewFileInfo = New FileInfo
    Call NewFileInfo.Init(FileName)
End Function

''
' Returns a new BinaryWriter used to write to a Stream object.
'
' @param Stream The stream to be written to by the writer.
' @param Encoding The encoding to be used when converting chars to bytes.
' @return The new BinaryWriter object.
' @remarks The Stream object must support writing.<br>
' <p>If no Encoding object is supplied, then an encoding of Cor.NewUTF8Encoding(False, True) is used.</p>
'
Public Function NewBinaryWriter(ByVal Stream As Stream, Optional ByVal Encoding As Encoding) As BinaryWriter
    Set NewBinaryWriter = New BinaryWriter
    Call NewBinaryWriter.Init(Stream, Encoding)
End Function

''
' Returns a new EndOfStreamException with the specified message and optional exception.
'
' @param Message A human-readable message that describes the cause of the exception.
' @param InnerException The exception that caused this exception.
' @return The new exception object.
' @remarks This exception can be thrown whenever an attempt to read passed the end of a stream
' is attempted. Such classes as BinaryReader may throw this exception if necessary.
'
Public Function NewEndOfStreamException(ByVal Message As String, Optional ByVal InnerException As Exception) As EndOfStreamException
    Set NewEndOfStreamException = New EndOfStreamException
    Call NewEndOfStreamException.Init(Message, InnerException)
End Function

''
' Returns a new BinaryReader that can read from the specified stream usind the specified Encoding.
'
' @param Stream The stream to read the data from.
' @param Encoding The encoding system to be used to decode the bytes from the stream.
' @return The new BinaryReader with the specified stream and encoding.
' @remarks If no Encoding is specified, then the reader uses an Encoding of New UTF8Encoding.
'
Public Function NewBinaryReader(ByVal Stream As Stream, Optional ByVal Encoding As Encoding) As BinaryReader
    Set NewBinaryReader = New BinaryReader
    Call NewBinaryReader.Init(Stream, Encoding)
End Function

''
' Returns a new OverflowException with the specified message and exception.
'
' @param Message A human-readable message describing the exception.
' @param InnerException The exception that caused this exception.
' @return The new OverflowException object.
' @return This exception is thrown when a value may not fit into an existing datatype.
'
Public Function NewOverflowException(ByVal Message As String, Optional ByVal InnerException As Exception) As OverflowException
    Set NewOverflowException = New OverflowException
    Call NewOverflowException.Init(Message, InnerException)
End Function

''
' Creates an enumerator used to iterate over the Unicode characters of a String.
'
' @param s The string to iterate each character of.
' @param CopyString Indicates if the enumerator should create a copy of the passed
' in string, or use the existing string in memory.
' @return A CharEnumerator used to iterate over each character of a String.
' @remarks By default <b>CharEnumerator</b> makes a copy of the passed in
' String. The original string can be changed as normal. If <i>CopyString</i> is set
' to False, then the original string must not change or go out of scope for the
' duration of the <b>CharEnumerator</b> object's existance. By telling <b>CharEnumerator</b>
' to not make a copy of the original string, memory is not allocated for the string,
' which can be helpful for large strings.
'
Public Function NewCharEnumerator(ByRef s As String, Optional ByVal CopyString As Boolean = True) As CharEnumerator
    Set NewCharEnumerator = New CharEnumerator
    Call NewCharEnumerator.Init(s, CopyString)
End Function

''
' Creates a new instance of the Gregorian calendar class.
'
' @param CalendarType The language type the calendar is set to.
' @return A new instance of a GregorianCalendar class.
'
Public Function NewGregorianCalendar(ByVal CalendarType As GregorianCalendarTypes) As GregorianCalendar
    Set NewGregorianCalendar = New GregorianCalendar
    Call NewGregorianCalendar.Init(CalendarType)
End Function

''
' Creates a new ApplicationException object.
'
' @param Message A human-readable message explaining the exception.
' @param InnerException An exception that caused this exception.
' @return A fully created ApplicationException object.
'
' @see ApplicationException
'
Public Function NewApplicationException(ByVal Message As String, Optional ByVal InnerException As Exception) As ApplicationException
    Set NewApplicationException = New ApplicationException
    Call NewApplicationException.Init(Message, InnerException)
End Function

''
' Enumerates through the resources in a .RES file.
'
' @param Source The filename or stream to read from.
' @return A reader used to enumerate the resources.
'
Public Function NewResourceReader(ByRef Source As Variant) As ResourceReader
    Set NewResourceReader = New ResourceReader
    Call NewResourceReader.Init(Source)
End Function

''
' Provides a method to create a new .RES file.
'
' @param Source A .RES filename or Stream.
' @return A ResourceWriter used to create a resource file.
'
Public Function NewResourceWriter(ByRef Source As Variant) As ResourceWriter
    Set NewResourceWriter = New ResourceWriter
    Call NewResourceWriter.Init(Source)
End Function

' for internal use only
Friend Function NewRegistryKey(ByVal Key As Long, ByVal Name As String, ByVal Writable As Boolean) As RegistryKey
    Set NewRegistryKey = New RegistryKey
    Call NewRegistryKey.Init(Key, Name, Writable)
End Function

''
' Allows for quickly creating a cDateTime object from a VB Date value.
'
' @param Value The date of the new cDateTime.
' @param Kind Specifies whether the cDateTime is a UTC, Local, or unspecified time type.
' @return A cDateTime object.
'
Public Function NewDate(ByVal Value As Date, Optional ByVal Kind As DateTimeKind = UnspecifiedKind) As cDateTime
    Set NewDate = modStaticClasses.cDateTime.FromOADate(Value, Kind)
End Function

''
' Creates a new set of resources for a specific culture.
'
' @param CultureID The specific culture to retrieve resources for.
' @param Source A .RES filename, IResourceReader, or Stream object.
' @return A new ResourceSet.
'
Public Function NewResourceSet(ByVal CultureID As Variant, ByRef Source As Variant) As ResourceSet
    Set NewResourceSet = New ResourceSet
    Call NewResourceSet.Init(CultureID, Source)
End Function

''
' Loads all resources from a source for all cultures in the source.
'
' @param Source A .RES file, IResourceReader or Stream to read all resources from.
' @return A new ResourceManager object.
' @remarks If a set of resource files is to be loaded into a resource manager
' then use the <b>ResourceManager.CreateFileBasedResourceManager</b> method.
'
' @see ResourceManager
' @see ResourceManagerStatic
'
Public Function NewResourceManager(ByRef Source As Variant) As ResourceManager
    Set NewResourceManager = New ResourceManager
    Call NewResourceManager.InitFromFile(Source)
End Function

''
' Creates a new Hashtable object with specific settings.
'
' @param d An IDictionary object containing keys and values to be added to the Hashtable.
' @param hcp A hashcode generator used in place of the default hashcode algorithm.
' @param Comparer A custom comparer used to compare specific values.
' @return A new Hashtable object.
'
Public Function NewHashtable(Optional ByVal d As IDictionary, Optional ByVal hcp As IHashcodeProvider, Optional ByVal Comparer As IComparer) As Hashtable
    Set NewHashtable = New Hashtable
    Call NewHashtable.InitDictionary(d, hcp, Comparer)
End Function

''
' Creates a new WinResourceReader reader object.
'
' @param FileName The path and name of the EXE or DLL from which to read the resources.
' @return A new WinResourceReader object.
'
Public Function NewWinResourceReader(ByVal FileName As String) As WinResourceReader
    Set NewWinResourceReader = New WinResourceReader
    Call NewWinResourceReader.Init(FileName)
End Function

''
' Creates a new DriveInfo object used to inspect a drive's information.
'
' @param Name The name of the drive (C:\, \\Server\Share).
' @return A <b>DriveInfo</b> object representing the specified drive.
'
Public Function NewDriveInfo(ByVal Name As String) As DriveInfo
    Set NewDriveInfo = New DriveInfo
    Call NewDriveInfo.Init(Name)
End Function

''
' Returns a new ConsoleKeyInfo object that represents the state of a key press.
'
' @param KeyChar The Asc value of the key pressed.
' @param Key The key code of the pressed key.
' @param Shift Whether the shift key was pressed.
' @param Alt Wheter the alt key was pressed.
' @param Control Whether the control key was pressed.
' @return A new ConsoleKeyInfo object.
'
Public Function NewConsoleKeyInfo(ByVal KeyChar As Integer, ByVal Key As ConsoleKey, ByVal Shift As Boolean, ByVal Alt As Boolean, ByVal Control As Boolean) As ConsoleKeyInfo
    Set NewConsoleKeyInfo = New ConsoleKeyInfo
    Call NewConsoleKeyInfo.Init(KeyChar, Key, Shift, Alt, Control)
End Function

''
' Returns a new TimeSpan representing the time portion of the date.
'
' @param DateValue The date to retrieve the time portion of.
' @return A TimeSpan object representing the time portion.
'
Public Function NewTime(ByVal DateValue As Variant) As TimeSpan
    Select Case VarType(DateValue)
        Case vbDate:    Set NewTime = NewTimeSpan(, Hour(DateValue), Minute(DateValue), Second(DateValue))
        Case Else:      Set NewTime = cDateTime.GetcDateTime(DateValue).TimeOfDay
    End Select
End Function

''
' Returns an object used to manipulate an INI file.
'
' @param FileName The INI filename.
' @return The INIFile object.
'
Public Function NewINIFile(ByVal FileName As String) As INIFile
    Set NewINIFile = New INIFile
    Call NewINIFile.Init(FileName)
End Function

''
' Creates a new Ticker.
'
' @param Interval The delay between events in milliseconds.
' @param Data A user supplied value to be associated with the timer.
' @param AutoReset <b>True</b> to have the timer event recur, <b>False</b> for one event.
' @param AddressOfCallback An address to a callback method that can be called when the timer has elapsed.
' Setting this to zero will stop the callback.
' @return A new Ticker object.
' @remarks A callback method is used to allow a timer object to
' notify the function without having to use an Event. A function callback
' must have the following signature, or the application may crash.
' <pre>
' Public Sub TickerCallback(ByRef Ticker As Ticker, ByRef Data As Variant)
'     ' do stuff
' End Sub
' </pre>
' The two parameters must be declared as ByRef.
'
Public Function NewTicker(ByVal Interval As Long, Optional ByVal Data As Variant, Optional ByVal AutoReset As Boolean = True, Optional ByVal AddressOfCallback As Long) As Ticker
    Set NewTicker = New Ticker
    Call NewTicker.Init(Interval, Data, AutoReset, AddressOfCallback)
End Function

''
' Creates an array of values.
'
' @param arraytype The target datatype of the resultant array and the values to be converted to.
' @param arrayelements The values to be put into the new array.
' @return An array of the specified datatype and containing the array elements with conversion,
' or an empty array if no elements are defined.
' <p>Elements of a size that are a multiple of 4 will not create a new array
' when being returned. All other datatype sizes will cause a new array
' to be created.</p>
'
Public Function NewArray(ByVal ArrayType As vbVarType, ParamArray ArrayElements() As Variant) As Variant
    Dim Args() As Variant
    Call Helper.Swap4(ByVal ArrPtr(Args), ByVal Helper.DerefEBP(16))
    NewArray = cArray.InternalNewArray(ArrayType, Args)
End Function

''
' Creates a new base utility class that can be contained within an exception class.
'
' @param Subclass The containing exception class that will contain this object instance.
' @param AppTitle The title of the application that contains the exception class.
' @param Message An understandable message that describes the exception being thrown.
' @param InnerException The exception that caused the current exception to be thrown.
' @param HResult The error code for the current exception.
' @param Source A description of the source of the current exception.
' @param HelpLink A file or URL link to a help resource.
' @return A new base class used for default handling of exception routines.
' @remarks Most exception classes simply contain this information to be retrieved for
' viewing. Few exception classes have additional requirements. This class can be used
' to handle the default set of information for an exception class. By containing this
' class inside the exception class, the exception class can delegate calls to this class
' for standard handling.
' <p>If additional values need to be maintained, then this class also supports the ability
' to set and get additional values using a String name as the key. The <b>ArgumentException</b>
' class uses this to hold the name of the parameter that caused the exception.</p>
' <p>This class is persistable and will also attempt to persist the custom values that
' have been added to it. If the persistence fails, then an error will occur.</p>
'
Public Function NewExceptionBase(Optional ByVal Subclass As Object, Optional ByVal AppTitle As String, Optional ByVal Message As String, Optional ByVal InnerException As Exception, Optional ByVal HResult As Long = 5, Optional ByVal Source As String, Optional ByVal HelpLink As String) As ExceptionBase
    Set NewExceptionBase = New ExceptionBase
    Call NewExceptionBase.Init(Subclass, AppTitle, Message, InnerException, HResult, Source, HelpLink)
End Function

''
' Creates a new Queue initialized with the elements in the collection.
'
' @param Col The elements to place in the queue upon creation.
' @return A Queue object with the elements placed in it.
' @remarks The <i>Col</i> parameter can be a VBA.Collection, ICollection object or an Array.
'
Public Function NewQueue(ByRef Col As Variant) As Queue
    Set NewQueue = New Queue
    Call NewQueue.Init(Col)
End Function

''
' Creates a new object used to uniquely identify a resource.
'
' @param Name The name of the resource being identified.
' @param ResourceType The type of resource being identified.
' @param LanguageID The Locale ID associated with the resource.
' @return An object that can be used to uniquely identify a resource.
' @remarks Name and ResourceType can be a string or number.
'
Public Function NewResourceKey(ByRef Name As Variant, ByRef ResourceType As Variant, ByVal LanguageID As Long) As ResourceKey
    Set NewResourceKey = New ResourceKey
    Call NewResourceKey.Init(Name, ResourceType, LanguageID)
End Function

''
' Creates a picture group object used to contain information
' about resources that are grouped together.
'
' @param GroupType The type of resource the group represents.
' @return A group of container for a specific resource type.
'
Public Function NewResPictureGroup(ByVal GroupType As PictureGroupTypes) As ResPictureGroup
    Set NewResPictureGroup = New ResPictureGroup
    Call NewResPictureGroup.InitNew(GroupType)
End Function

''
' Creates a new ExternalException object.
'
' @param Message A user-defined message specific to the exception.
' @param InnerException A reference to an exception that caused this exception.
' @return A new ExternalException object.
'
Public Function NewExternalException(ByVal Message As String, Optional ByVal InnerException As Exception) As ExternalException
    Set NewExternalException = New ExternalException
    Call NewExternalException.Init(Message, InnerException)
End Function

''
' Creates a new CspParameters object.
'
' @param TypeIn The type code used to specify the kind of provider to create.
' @param ProviderNameIn The provider name.
' @param ContainerNameIn The provider container name.
' @return Returns a new CspParameters object.
'
Public Function NewCspParameters(ByVal TypeIn As Long, Optional ByVal ProviderNameIn As String, Optional ByVal ContainerNameIn As String) As CspParameters
    Set NewCspParameters = New CspParameters
    Call NewCspParameters.Init(TypeIn, ProviderNameIn, ContainerNameIn)
End Function

''
' Creates a new random number generator that uses a crypto service provider to generate random numbers.
'
' @param CspParams Parameters used to allow a specific crypto service provider to be used instead of the default.
' @return A new RNGCryptoServiceProvider object that generated random numbers using a crypto service provider.
' @remarks Crypto service providers generally produce a more random set of numbers than the standard
' Windows API set of random number generators. This is because cryptography depends a great deal on
' go random number generation to be effective.
' <p>The default provider will be one of Microsoft's cryptography service providers such as
' "Microsoft Base Cryptographic Provider v1.0" or "Microsoft Enhanced Cryptographic Provider v1.0."</p>
' <p>Using the <b>CspParams</b> allows for a specific or custom provider to be used instead of the
' Microsoft providers.</p>
' @see RNGCryptoServiceProvider
'
Public Function NewRNGCryptoServiceProvider(ByVal CspParams As CspParameters) As RNGCryptoServiceProvider
    Set NewRNGCryptoServiceProvider = New RNGCryptoServiceProvider
    Call NewRNGCryptoServiceProvider.Init(CspParams)
End Function

''
' Creates a new CryptographicException object.
'
' @param Message The message contained within the exception object.
' @param InnerException A reference to an exception that may have cause this exception to occur.
' @return A new CryptoGraphicException object.
' @remarks Most of the exceptions that occur within the cryptography set of classes will be
' a CryptographicException.
' @see CryptographicException
'
Public Function NewCryptographicException(ByVal Message As String, Optional ByVal InnerException As Exception) As CryptographicException
    Set NewCryptographicException = New CryptographicException
    Call NewCryptographicException.Init(Message, InnerException)
End Function

''
' Creates a new KeySizes object with the Min, Max and Skip sizes defined.
'
' @param MinSize The smallest key size defined by this KeySizes instance.
' @param MaxSize The largest key size defined by this KeySizes instance.
' @param SkipSize The step size from MinSize to MaxSize of legal key sizes
' between the smallest and largest legal key sizes in this instance.
' @return A new KeySizes object with the Min, Max and Skip sizes defined.
' @remarks A KeySizes object represents a set of legal key sizes. It defines all the possible
' legal key sizes by defining the smallest and largest key sizes allows, then defining a step
' size used to calculate the additional legal key sizes between the smallest and largest key sizes.
' <p>If the smallest key size is 64 and the largest key size is 128 and a skip size of 32 is defined,
' then all possible legal key sizes are from 64 to 128 in increments of 32. In this case, the valid
' key sizes are 64, 96 and 128.</p>
' @see KeySizes
'
Public Function NewKeySizes(ByVal MinSize As Long, ByVal MaxSize As Long, ByVal SkipSize As Long) As KeySizes
    Set NewKeySizes = New KeySizes
    Call NewKeySizes.Init(MinSize, MaxSize, SkipSize)
End Function

''
' Creates a new FromBase64Transform object used to convert base64 characters to plain text.
'
' @param WhiteSpaces Determines how whitespace should be handled during the transformation process.
' @return A new FromBase64Transfrom object.
' @see FromBase64Transform
'
Public Function NewFromBase64Transform(ByVal WhiteSpaces As FromBase64TransformMode) As FromBase64Transform
    Set NewFromBase64Transform = New FromBase64Transform
    Call NewFromBase64Transform.Init(WhiteSpaces)
End Function

''
' Creates a new CryptoStream object used to read or write transformed data to or from an underlying stream.
'
' @param Stream The underlying stream to read or write the transformed data.
' @param Transform The object used to transfrom data to or from plain text.
' @param Mode Sets the CryptoStream to be either in read mode or write mode.
' @return A new CryptoStream object.
' @see CryptoStream
'
Public Function NewCryptoStream(ByVal Stream As Stream, ByVal Transform As ICryptoTransform, ByVal Mode As CryptoStreamMode) As CryptoStream
    Set NewCryptoStream = New CryptoStream
    Call NewCryptoStream.Init(Stream, Transform, Mode)
End Function

''
' Returns an array of Longs generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a Long or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewLongs(ParamArray Values() As Variant) As Long()
    Dim Ret()   As Long
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbLong, 0, 0)
    End If
    
    NewLongs = Ret
End Function

''
' Returns an array of Bytes generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a byte or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewBytes(ParamArray Values() As Variant) As Byte()
    Dim Ret()   As Byte
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbByte, 0, 0)
    End If
    
    NewBytes = Ret
End Function

''
' Returns an array of Integers generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to an Integer or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewIntegers(ParamArray Values() As Variant) As Integer()
    Dim Ret()   As Integer
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbInteger, 0, 0)
    End If
    
    NewIntegers = Ret
End Function

''
' Returns an array of Strings generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a String or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewStrings(ParamArray Values() As Variant) As String()
    Dim Ret()   As String
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbString, 0, 0)
    End If
    
    NewStrings = Ret
End Function

''
' Returns an array of Singles generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a Single or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewSingles(ParamArray Values() As Variant) As Single()
    Dim Ret()   As Single
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbSingle, 0, 0)
    End If
    
    NewSingles = Ret
End Function

''
' Returns an array of Doubles generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a Double or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewDoubles(ParamArray Values() As Variant) As Double()
    Dim Ret()   As Double
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbDouble, 0, 0)
    End If
    
    NewDoubles = Ret
End Function

Public Function NewDecimals(ParamArray Values() As Variant) As Variant()
    Dim Result()    As Variant
    Dim ub          As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Result(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Result(i) = CDec(Values(i))
        Next
    Else
        SAPtr(Result) = SafeArrayCreateVector(vbVariant, 0, 0)
    End If
    
    NewDecimals = Result
End Function

''
' Returns an array of Date values generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to dates or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewDates(ParamArray Values() As Variant) As Date()
    Dim Ret()   As Date
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbDate, 0, 0)
    End If
    
    NewDates = Ret
End Function

''
' Returns an array of Currency values generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to Currency or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewCurrencys(ParamArray Values() As Variant) As Currency()
    Dim Ret()   As Currency
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbCurrency, 0, 0)
    End If
    
    NewCurrencys = Ret
End Function

''
' Returns an array of Boolean values generated from the values passed in.
'
' @param Values The values to create the array from.
' @remarks The values need to be convertable to a boolean or an exception will occur.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewBooleans(ParamArray Values() As Variant) As Boolean()
    Dim Ret()   As Boolean
    Dim ub      As Long
    
    ub = UBound(Values)
    If ub >= 0 Then
        ReDim Ret(0 To ub)
        Dim i As Long
        For i = 0 To ub
            Ret(i) = Values(i)
        Next i
    Else
        SAPtr(Ret) = SafeArrayCreateVector(vbBoolean, 0, 0)
    End If
    
    NewBooleans = Ret
End Function

''
' Returns a Variant array containing the values passed in.
'
' @param Values The values to create the array from.
' @remarks This is the same as using the Array function.
' <p>The array is zero-based.</p>
' <p>If no values are passed in then an empty array is returned, not a null array.</p>
'
Public Function NewVariants(ParamArray Values() As Variant) As Variant()
    NewVariants = Values
End Function

''
' Creates a new CspKeyContainerInfo object.
'
' @param Parameters The parameters used to derive information from.
' @return A new CspKeyContainerInfo object.
' @remarks This object is usually created by crypto service provider. Instantiating an
' object directly may result in property throwing an exception when accessed.
'
Public Function NewCspKeyContainerInfo(ByVal Parameters As CspParameters) As CspKeyContainerInfo
    Set NewCspKeyContainerInfo = New CspKeyContainerInfo
    Call NewCspKeyContainerInfo.Init(Parameters)
End Function

''
' Creates a new Rfc2898DeriveBytes object.
'
' @param Password This can be a String password or a Byte array.
' @param Salt This can be a Byte array to be used as salt or an Integer value indicating the size of the salt to be generated.
' @param Iterations The number of times the operation should be performed when generating a password.
' @return A new Rfc2898DeriveBytes object.
' @remarks This is used to create Byte array passwords from standard plain text passwords.
'
Public Function NewRfc2898DeriveBytes(ByRef Password As Variant, ByRef Salt As Variant, Optional ByVal Iterations As Long = 1000) As Rfc2898DeriveBytes
    Set NewRfc2898DeriveBytes = New Rfc2898DeriveBytes
    Call NewRfc2898DeriveBytes.Init(Password, Salt, Iterations)
End Function

''
' Creates a base classed used by hash algorithms.
'
' @param Callback Used to call back into the specifics of an algorithm implementation.
' @param KeepWeakReference Indicates if the base class should maintain a weak reference to the
' Callback object to prevent a circular reference from occuring if the Callback object is the same
' object as the one instantiating this instance.
' @param UseBuffer Indicates if the internal buffering mechanism in the base class should be used,
' or if all data should be passed to the Callback object to allow custom buffering.
' @param BlockSize The size of the byte block that is processed at a time by the specific algorithm
' using this base class. This is to help with the buffering.
' @return A new HashAlgorithmBase object.
' @remarks This base class can handle most hash algorithm implementation requirement except for the
' specifics of a hash algorithm. That is where the Callback is used.
' @see HashAlgorithmBase
'
Public Function NewHashAlgorithmBase(ByVal Callback As IHashAlgorithm, Optional ByVal KeepWeakReference As Boolean, Optional ByVal UseBuffer As Boolean, Optional ByVal BlockSize As Long) As HashAlgorithmBase
    Set NewHashAlgorithmBase = New HashAlgorithmBase
    Call NewHashAlgorithmBase.Init(Callback, KeepWeakReference, UseBuffer, BlockSize)
End Function

''
' Creates a new CryptoAPIHash object.
'
' @param HashAlg The algorithm of the hash to be used.
' @param HashSize The size of the hash output.
' @return A new CryptoAPIHash object.
' @remarks A CryptoAPIHash object is an implementation of the IHashAlgorithm callback that
' uses the CryptoAPI hash routines.
' @see CryptoAPIHash
'
Public Function NewCryptoAPIHash(ByVal HashAlg As Long, ByVal HashSize As Long) As CryptoAPIHash
    Set NewCryptoAPIHash = New CryptoAPIHash
    Call NewCryptoAPIHash.Init(HashAlg, HashSize)
End Function

''
' Creates a new HMACSHA1 object.
'
' @param Key The key used to sign the message.
' @param UseManagedSha1 Indicates if the SHA1Managed or SHA1CryptoServiceProvider should be used.
' @return A new HMACSHA1 object.
' @see HMACSHA1
'
Public Function NewHMACSHA1(ByRef Key() As Byte, Optional ByVal UseManagedSha1 As Boolean = False) As HMACSHA1
    Set NewHMACSHA1 = New HMACSHA1
    Call NewHMACSHA1.Init(Key, UseManagedSha1)
End Function

''
' Creates a base class that can be used as a base implementation of an HMAC class.
'
' @param HashAlgorithm An instance of an algorithm to be used Hash Message Authentication Code.
' @param KeySize The size of the key to be used for signing the message.
' @param HashName The name of the hash being used.
' @param Key The key used to sign the message.
' @return A new HMACBase object.
' @remarks This class implements many of the common routines used by HMAC classes.
'
Public Function NewHMACBase(ByVal HashAlgorithm As HashAlgorithm, ByVal KeySize As Long, Optional ByVal HashName As String, Optional ByRef Key As Variant) As HMACBase
    Set NewHMACBase = New HMACBase
    Call NewHMACBase.Init(HashAlgorithm, KeySize, HashName, Key)
End Function

''
' Creates a new HMACSHA256 object.
'
' @param Key The key used to sign the message.
' @return A new HMACSHA256 object.
' @see HMACSHA256
'
Public Function NewHMACSHA256(ByRef Key() As Byte) As HMACSHA256
    Set NewHMACSHA256 = New HMACSHA256
    Call NewHMACSHA256.Init(Key)
End Function

''
' Creates a new HMACSHA384 object.
'
' @param Key The key used to sign the message.
' @return A new HMACSHA384 object.
' @see HMACSHA384
'
Public Function NewHMACSHA384(ByRef Key() As Byte) As HMACSHA384
    Set NewHMACSHA384 = New HMACSHA384
    Call NewHMACSHA384.Init(Key)
End Function

''
' Creates a new HMACSHA512 object.
'
' @param Key The key used to sign the message.
' @return A new HMACSHA512 object.
' @see HMACSHA512
'
Public Function NewHMACSHA512(ByRef Key() As Byte) As HMACSHA512
    Set NewHMACSHA512 = New HMACSHA512
    Call NewHMACSHA512.Init(Key)
End Function

''
' Creates a new HMACMD5 object.
'
' @param Key The key used to sign the message.
' @return A new HMACMD5 object.
' @see HMACMD5
'
Public Function NewHMACMD5(ByRef Key() As Byte) As HMACMD5
    Set NewHMACMD5 = New HMACMD5
    Call NewHMACMD5.Init(Key)
End Function

''
' Creates a new HMACRIPEMD160 object.
'
' @param Key The key used to sign the message.
' @return A new HMACRIPEMD160 object.
' @see HMACRIPEMD160
'
Public Function NewHMACRIPEMD160(ByRef Key() As Byte) As HMACRIPEMD160
    Set NewHMACRIPEMD160 = New HMACRIPEMD160
    Call NewHMACRIPEMD160.Init(Key)
End Function

''
' Creates a new RSACryptoServiceProvider object.
'
' @param KeySize The size of the key used by the service provider.
' @param Parameters A CspParameters object used to specify custom service provider options.
' @return A new RSACryptoServiceProvider.
' @remarks The key size can is 384bit to 512bit for non-high encryption machines and 384bit to 16384bit for high encryption machines.
' @see RSACryptoServiceProvider
'
Public Function NewRSACryptoServiceProvider(Optional ByVal KeySize As Long, Optional ByVal Parameters As CspParameters) As RSACryptoServiceProvider
    Set NewRSACryptoServiceProvider = New RSACryptoServiceProvider
    Call NewRSACryptoServiceProvider.Init(KeySize, Parameters)
End Function

''
' Creates a new DSACryptoServiceProvider object.
'
' @param KeySize The key size in bits. This can be from 384bits to 1024bits.
' @param Parameters A CspParameters object used to specify custom service provider options.
' @return A new DSACryptoServiceProvider object.
' @see DSACryptoServiceProvider
'
Public Function NewDSACryptoServiceProvider(Optional ByVal KeySize As Long, Optional ByVal Parameters As CspParameters) As DSACryptoServiceProvider
    Set NewDSACryptoServiceProvider = New DSACryptoServiceProvider
    Call NewDSACryptoServiceProvider.Init(KeySize, Parameters)
End Function

''
' Creates a new SecurityElement object.
'
' @param Tag The name of the security element.
' @param Text The text within the element tags.
' @return A new SecurityElement object.
' @remarks A SecurityElement is a simple version of an XML tag with attributes and text.
' @see SecurityElement
'
Public Function NewSecurityElement(ByVal Tag As String, Optional ByVal Text As String) As SecurityElement
    Set NewSecurityElement = New SecurityElement
    Call NewSecurityElement.Init(Tag, Text)
End Function

''
' Creates a new XmlSyntaxException object.
'
' @param LineNumber The number in the Xml string the exception occurred.
' @param Message The exception message.
' @param InnerException The exception that caused this exception to occur.
' @return A new XmlSyntaxException object.
' @see XmlSyntaxException
'
Public Function NewXmlSyntaxException(Optional ByVal LineNumber As Variant, Optional ByVal Message As Variant, Optional ByVal InnerException As Exception) As XmlSyntaxException
    Set NewXmlSyntaxException = New XmlSyntaxException
    Call NewXmlSyntaxException.Init(LineNumber, Message, InnerException)
End Function

''
' Creates a new MACTripleDES object.
'
' @param RgbKey The key to be used by the implementation of the TripleDES algorithm.
' @param StrTripleDES The name of the TripleDES implementation to be used.
' @return A new MACTripleDES object.
'
Public Function NewMACTripleDES(ByRef RgbKey() As Byte, Optional ByVal StrTripleDES As String = "3DES") As MACTripleDES
    Set NewMACTripleDES = New MACTripleDES
    Call NewMACTripleDES.Init(RgbKey, StrTripleDES)
End Function

''
' Creates a new BigInteger instance.
'
' @param value The value to intialize a new instance of <b>BitInteger</b> to.
' @return A new BigInteger instance containing the specified value.
' @remarks The follow datatypes are supported:<br>
' vbLong<br>
' vbInteger<br>
' vbByte<br>
' vbDoube<br>
' vbSingle<br>
' vbCurrency<br>
' vbDecimal<br>
' vbByte Or vbArray<br>
'
Public Function NewBigInteger(ByRef Value As Variant) As BigInteger
    Set NewBigInteger = New BigInteger
    Call NewBigInteger.Init(Value)
End Function

''
' Attempts to convert a value to a valid BigInteger instance.
'
' @param value The value to be converted to BigInteger.
' @return A new BigInteger instance representing the converted value.
' @remarks The follow datatypes are supported:<br>
' vbString<br>
' vbLong<br>
' vbInteger<br>
' vbByte<br>
' vbDoube<br>
' vbSingle<br>
' vbCurrency<br>
' vbDecimal<br>
' vbByte Or vbArray<br>
' BigInteger (returns same instance)<br>
' Nothing (returns BigInteger.Zero)<br>
'
' <pre>
' Dim b As BigInteger
' Set b = BInt(<i>somevalue</i>)
' </pre>
'
Public Function BInt(ByRef Value As Variant) As BigInteger
    Select Case VarType(Value)
        Case vbObject, vbDataObject
            If Value Is Nothing Then
                Set BInt = BigInteger.Zero
            Else
                Set BInt = Value
            End If
            
        Case vbString
            Set BInt = BigInteger.Parse(Value)
            
        Case Else
            Set BInt = NewBigInteger(Value)
    End Select
End Function


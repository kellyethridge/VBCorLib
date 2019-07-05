VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: IniFile
'

''
' Provides methods for manipulating an INI initialization file.
'
' @remarks An INI file contains sections in a format like<br><br>
' <i>
' [section]<br>
' key=value<br><br>
' </i>
' A new IniFile object can be created using the <b>NewIniFile</b> method.
' <pre>
' Dim ini As IniFile
' Set ini = NewIniFile("c:\settings.ini")
' </pre>
' <p>If the file name does not specify the directory, then the file will
' be created in the Windows directory. In order to specify the current
' directory, preceed the file name with a current directory symbol ".".
' <pre>
' Dim Ini As IniFile
' Set Ini = NewIniFile(".\local.ini")
' </pre>
' </p>
'
' @see IniResourceWriter
'
Option Explicit
Implements IObject

Private mFileName   As String
Private mAutoFlush  As Boolean


''
' Returns the auto-flush state of the object.
'
' @Return <b>True</> if <b>AutoFlush</b> is active, <b>False</b> otherwise.
' @remarks When <b>AutoFlush</b> is set to <b>True</b>, the object will call <a href="IniFile.Flush.html">Flush</a> after any method that writes to the INI file.
Public Property Get AutoFlush() As Boolean
    AutoFlush = mAutoFlush
End Property

''
' Sets the auto-flush state of the object.
'
' @param Value The new auto-flush state of the object.
' @remarks When <b>AutoFlush</b> is set to <b>True</b>, the object will call <a href="IniFile.Flush.html">Flush</a> after any method that writes to the INI file.
Public Property Let AutoFlush(ByVal Value As Boolean)
    mAutoFlush = Value
End Property

''
' Flushes any cached updates to the INI file.
Public Sub Flush()
    WritePrivateProfileStringW vbNullString, vbNullString, vbNullString, vbNullString
End Sub

''
' Returns the name of the INI file being manipulated.
'
' @return The INI file name.
Public Property Get FileName() As String
    FileName = mFileName
End Property

''
' Sets a key-value pair within a section of an INI file.
'
' @param Section The INI section within the file to set the value.
' @param Key The name of the value to be set.
' @param Value The value to be set in the INI file.
' @remarks If an object is passed in, then it must implement the <a href="IObject.html">IObject</a>
' interface or an exception will be thrown. All other datatypes will be converted to their normal string value.
Public Sub SetValue(ByRef Section As String, ByRef Key As String, ByRef Value As Variant)
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    If Len(Key) = 0 Then _
        Error.Argument Argument_EmptyIniKey
    
    Dim s As String
    ' We don't want a null string, since that would
    ' delete the INI entry instead.
    s = IfString(Convert.ToString(Value), "")

    If WritePrivateProfileStringW(Section, Key, s, mFileName) = BOOL_FALSE Then _
        IOError.WinIOError Err.LastDllError
    
    If mAutoFlush Then
        Flush
    End If
End Sub

''
' Returns the string value of the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section.
' <p>An empty string is a valid value in an INI file. Testing for an empty string does not mean the value was not found.</p>
Public Function GetString(ByRef Section As String, ByRef Key As String, Optional ByVal Default As String) As String
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    If Len(Key) = 0 Then _
        Error.Argument Argument_EmptyIniKey

    ' MSDN says we cannot pass in a null value for the default.
    Default = IfString(Default, "")
    
    Dim Size    As Long
    Dim Buf     As String
    
    Size = 512
    
    Do
        Size = Size * 2
        Buf = String$(Size, vbNullChar)
        Size = GetPrivateProfileStringW(Section, Key, Default, Buf, Size, mFileName)
    Loop While Size = Len(Buf) - 1
    
    GetString = Left$(Buf, Size)
End Function

''
' Returns a <b>Long</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Long</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Long</b> using the <b>CLng</b> function. If an error occurs, then the default is returned.
Public Function GetInt32(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Long) As Long
    On Error GoTo Catch
    GetInt32 = CLng(GetString(Section, Key))
    Exit Function
    
Catch:
    GetInt32 = Default
End Function

''
' Returns a <b>Integer</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to an <b>Integer</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into an <b>Integer</b> using the <b>CInt</b> function. If an error occurs, then the default is returned.
Public Function GetInt16(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Integer) As Integer
    On Error GoTo Catch
    GetInt16 = CInt(GetString(Section, Key))
    Exit Function
    
Catch:
    GetInt16 = Default
End Function

''
' Returns a <b>Byte</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Byte</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Byte</b> using the <b>CByte</b> function. If an error occurs, then the default is returned.
Public Function GetByte(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Byte) As Byte
    On Error GoTo Catch
    GetByte = CByte(GetString(Section, Key))
    Exit Function
    
Catch:
    GetByte = Default
End Function

''
' Returns a <b>Boolean</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Boolean</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Boolean</b> using the <b>CBool</b> function. If an error occurs, then the default is returned.
Public Function GetBoolean(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Boolean) As Boolean
    On Error GoTo Catch
    GetBoolean = CBool(GetString(Section, Key))
    Exit Function
    
Catch:
    GetBoolean = Default
End Function

''
' Returns a <b>Date</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Date</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Date</b> using the <b>CDate</b> function. If an error occurs, then the default is returned.
Public Function GetDate(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Date) As Date
    On Error GoTo Catch
    GetDate = CDate(GetString(Section, Key))
    Exit Function
    
Catch:
    GetDate = Default
End Function

''
' Returns a <b>Double</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Double</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Double</b> using the <b>CDbl</b> function. If an error occurs, then the default is returned.
Public Function GetDouble(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Double) As Double
    On Error GoTo Catch
    GetDouble = CDbl(GetString(Section, Key))
    Exit Function
    
Catch:
    GetDouble = Default
End Function

''
' Returns a <b>Single</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Single</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Single</b> using the <b>CSng</b> function. If an error occurs, then the default is returned.
Public Function GetSingle(ByRef Section As String, ByRef Key As String, Optional ByVal Default As Single) As Single
    On Error GoTo Catch
    GetSingle = CSng(GetString(Section, Key))
    Exit Function
    
Catch:
    GetSingle = Default
End Function

''
' Returns a <b>Currency</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Currency</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Currency</b> using the <b>CCur</b> function. If an error occurs, then the default is returned.
Public Function GetCurrency(ByVal Section As String, ByVal Key As String, Optional ByVal Default As Currency) As Currency
    On Error GoTo Catch
    GetCurrency = CCur(GetString(Section, Key))
    Exit Function
    
Catch:
    GetCurrency = Default
End Function

''
' Returns a <b>Decimal</b> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <b>Decimal</b>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <b>Decimal</b> using the <b>CDec</b> function. If an error occurs, then the default is returned.
Public Function GetDecimal(ByRef Section As String, ByRef Key As String, Optional ByRef Default As Variant) As Variant
    Dim DecimalDefault As Variant
    
    If IsMissing(Default) Then
        DecimalDefault = CDec(0)
    Else
        DecimalDefault = CDec(Default)
    End If
    
    On Error GoTo Catch
    GetDecimal = CDec(GetString(Section, Key))
    Exit Function
    
Catch:
    GetDecimal = DecimalDefault
End Function

''
' Returns a <a href="CorDateTime.html">CorDateTime</a> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <a href="CorDateTime.html">CorDateTime</a>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <a href="CorDateTime.html">CorDateTime</a> using the <a href="CorDateTimeStatic.TryParse.html">CorDateTime.TryParse</a> function. If an error occurs, then the default is returned.
' </p>Passing in <b>Nothing</b> as the default will return <a href="CorDateTimeStatic.get.MinValue.html">CorDateTime.MinValue</a>.</p>
Public Function GetDateTime(ByRef Section As String, ByRef Key As String, Optional ByVal Default As CorDateTime) As CorDateTime
    Dim s As String
    s = GetString(Section, Key)
    
    If Not CorDateTime.TryParse(s, GetDateTime) Then
        Set GetDateTime = IfObject(Default, CorDateTime.MinValue)
    End If
End Function

''
' Returns a <a href="TimeSpan.html">TimeSpan</a> value from the specified key in the specified section of an INI file.
'
' @param Section The section within the INI file to search for the key.
' @param Key The key in the section of an INI file to retrieve the value of.
' @param Default The default value to return if the key is not found in the section, or the value could not be converted to a <a href="TimeSpan.html">TimeSpan</a>.
' @remarks An INI file contains all values as strings. The value is converted back
' into a <a href="TimeSpan.html">TimeSpan</a> by first attempting to parse the string as a <a href="TimeSpan.html">TimeSpan</a> object. If that
' fails, an attempt is made by testing for a <b>Date</b> value and retrieving the time portion. If an error occurs, then the default is returned.
' <p>Passing in <b>Nothing</b> for the default will return <a href="TimeSpanStatic.get.Zero.html">TimeSpan.Zero</b> for the default.</p>
Public Function GetTimeSpan(ByRef Section As String, ByRef Key As String, Optional ByVal Default As TimeSpan) As TimeSpan
    Dim s As String
    s = GetString(Section, Key)
    
    If Not TimeSpan.TryParse(s, GetTimeSpan) Then
        If IsDate(s) Then
            Set GetTimeSpan = TimeSpan.FromDate(s)
        Else
            Set GetTimeSpan = IfObject(Default, TimeSpan.Zero)
        End If
    End If
End Function

''
' Returns a list of section names in an INI file.
'
' @return An array of section names, or a zero-length array if no sections exist.
Public Function GetSectionNames() As String()
    GetSectionNames = GetList(vbNullString)
End Function

''
' Returns a list of key names in an INI file.
'
' @param Section The section within an INI file to retrieve key names.
' @return An array of key names from a section, or a zero-length array if no keys exist in the section.
Public Function GetKeyNames(ByRef Section As String) As String()
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection

    GetKeyNames = GetList(Section)
End Function

''
' Deletes a key-value pair in a specific section within an INI file.
'
' @param Section The section the key-value pair is to be deleted from.
' @param Key The key of the key-value pair to be deleted.
Public Sub DeleteKey(ByRef Section As String, ByRef Key As String)
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    If Len(Key) = 0 Then _
        Error.Argument Argument_EmptyIniKey
    
    If WritePrivateProfileStringW(Section, Key, vbNullString, mFileName) = BOOL_FALSE Then _
        IOError.WinIOError Err.LastDllError
    
    If mAutoFlush Then
        Flush
    End If
End Sub

''
' Deletes an entire section from an INI file, including all key-value pairs.
'
' @param Section The name of the section to be deleted from an INI file.
Public Sub DeleteSection(ByRef Section As String)
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    
    If WritePrivateProfileStringW(Section, vbNullString, vbNullString, mFileName) = BOOL_FALSE Then _
        IOError.WinIOError Err.LastDllError
    
    If mAutoFlush Then
        Flush
    End If
End Sub

''
' Returns an <a href="IDictionary.html">IDictionary</a> object containing the key-value pairs from a section in an INI file.
'
' @param Section The section to retrieve all the key-value pairs from.
' @return An <a href="IDictionary.html">IDictionary</a> object containing zero or more key-value pairs.
Public Function GetValues(ByRef Section As String) As IDictionary
    If Len(Section) = 0 Then _
        Error.Argument Argument_EmptyIniSection
    
    Dim Buf     As String
    Dim Size    As Long
    Dim Values  As Hashtable
    
    Size = 16384
    
    Do
        Size = Size * 2
        Buf = String$(Size, vbNullChar)
        Size = GetPrivateProfileSectionW(Section, Buf, Size, mFileName)
    Loop While Size = Len(Buf) - 2
    
    Set Values = Cor.NewHashtable(EqualityComparer:=StringComparer.OrdinalIgnoreCase)
    
    If Size > 0 Then
        Dim Entries() As String
        Entries = Split(Left$(Buf, Size - 1), vbNullChar)
    
        Dim KeyValue()  As String
        Dim i           As Long
        For i = 0 To UBound(Entries)
            KeyValue = Split(Entries(i), "=", 2)
            
            If UBound(KeyValue) = 1 Then
                Values(KeyValue(0)) = Trim$(KeyValue(1))
            Else
                Values(KeyValue(0)) = ""
            End If
        Next i
    End If
    
    Set GetValues = Values
End Function

''
' Returns a <a href="IniResourceWriter.html">IniResourceWriter</a> used to create an entire section withing an INI file at once.
'
' @param Section The section to create using the writer.
' @return The <a href="IniResourceWriter.html">IniResourceWriter</a> used to create an entire section.
' @see IniResourceWriter
Public Function GetSectionWriter(ByRef Section As String) As IniResourceWriter
    Set GetSectionWriter = Cor.NewIniResourceWriter(mFileName, Section)
End Function

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = MyBase.ToString(Me, App)
End Function

''
' Returns a boolean indicating if the value and this object instance are the same instance.
'
' @param value The value to compare equalit to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef FileName As String)
    If Len(FileName) = 0 Then _
        Error.Argument Argument_EmptyPath
    
    mFileName = FileName
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetList(ByRef Section As String) As String()
    Dim Size    As Long
    Dim Buf     As String
    Dim List()  As String
    
    Size = 512
    
    Do
        Size = Size * 2
        Buf = String$(Size, vbNullChar)
        Size = GetPrivateProfileStringW(Section, vbNullString, "", Buf, Size, mFileName)
    Loop While Size = Len(Buf) - 1
    
    If Size > 0 Then
        List = Split(CorString.TrimEnd(Left$(Buf, Size - 1), vbNullChar), vbNullChar)
    Else
        List = Cor.NewStrings()
    End If
    
    GetList = List
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfo"
Attribute VB_GlobalNameSpace = False
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
' Module: CultureInfo
'

''
' Represents information about a culture, such as number and date formatting.
'
' @see Constructors
' @see CultureInfoStatic
' @see ICloneable
' @see IFormatProvider
'
Option Explicit
Implements IObject
Implements ICloneable
Implements IFormatProvider

Private Const PROP_LCID                             As String = "LCID"
Private Const PROP_PARENTLCID                       As String = "ParentLCID"
Private Const PROP_NAME                             As String = "Name"
Private Const PROP_ENGLISHNAME                      As String = "EnglishName"
Private Const PROP_DISPLAYNAME                      As String = "DisplayName"
Private Const PROP_NATIVENAME                       As String = "NativeName"
Private Const PROP_THREELETTERISOLANGUAGENAME       As String = "ThreeLetterISOLanguageName"
Private Const PROP_THREELETTERWINDOWSLANGUAGENAME   As String = "ThreeLetterWindowsLanguageName"
Private Const PROP_TWOLETTERISOLANGUAGENAME         As String = "TwoLetterISOLanguageName"
Private Const PROP_USEUSEROVERRIDE                  As String = "UseUserOverride"
Private Const PROP_ISREADONLY                       As String = "IsReadOnly"
Private Const PROP_DATETIMEFORMAT                   As String = "DateTimeFormat"
Private Const PROP_NUMBERFORMAT                     As String = "NumberFormat"
Private Const PROP_CALENDAR                         As String = "Calendar"

Private Const CAL_GREGORIAN                 As Long = 1         ' Gregorian (localized) calendar
Private Const CAL_GREGORIAN_US              As Long = 2         ' Gregorian (U.S.) calendar
Private Const CAL_JAPAN                     As Long = 3         ' Japanese Emperor Era calendar
Private Const CAL_TAIWAN                    As Long = 4         ' Taiwan Era calendar
Private Const CAL_KOREA                     As Long = 5         ' Korean Tangun Era calendar
Private Const CAL_HIJRI                     As Long = 6         ' Hijri (Arabic Lunar) calendar
Private Const CAL_THAI                      As Long = 7         ' Thai calendar
Private Const CAL_HEBREW                    As Long = 8         ' Hebrew (Lunar) calendar
Private Const CAL_GREGORIAN_ME_FRENCH       As Long = 9         ' Gregorian Middle East French calendar
Private Const CAL_GREGORIAN_ARABIC          As Long = 10        ' Gregorian Arabic calendar
Private Const CAL_GREGORIAN_XLIT_ENGLISH    As Long = 11        ' Gregorian Transliterated English calendar
Private Const CAL_GREGORIAN_XLIT_FRENCH     As Long = 12        ' Gregorian Transliterated French calendar
Private Const CAL_JULIAN                    As Long = 13        ' Julian calendar


Private Type PropsType
    LCID                            As Long
    ParentLCID                      As Long
    Name                            As String
    EnglishName                     As String
    DisplayName                     As String
    NativeName                      As String
    ThreeLetterISOLanguageName      As String
    ThreeLetterWindowLanguageName   As String
    TwoLetterISOLanguageName        As String
    UseUserOverride                 As Boolean
    IsReadOnly                      As Boolean
End Type

Private mProps                  As PropsType
Private mNumberFormat           As NumberFormatInfo
Private mDateTimeFormat         As DateTimeFormatInfo
Private mLoaded                 As Boolean
Private mCalendar               As Calendar
Private mOptionalCalendars()    As Calendar


''
' Returns the default calendar for the current culture.
'
' @return The calendar associated with the current culture.
'
Public Property Get Calendar() As Calendar
    If mCalendar Is Nothing Then
        Set mCalendar = GetCalendar(CultureTable.GetNumber(mProps.LCID, ICALENDARTYPE))
    End If
    
    Set Calendar = mCalendar
End Property

''
' Returns a list of optional calendars supported by the culture.
'
' @return A list of calendars associated with the current culture.
'
Public Property Get OptionalCalendars() As Calendar()
    If CorArray.IsNull(mOptionalCalendars) Then
        Dim CalendarTypes() As Long
        Dim i As Long
        
        CalendarTypes = CultureTable.GetNumberArray(mProps.LCID, SOPTIONALCALENDARS)
        ReDim mOptionalCalendars(0 To UBound(CalendarTypes))
        
        For i = 0 To UBound(CalendarTypes)
            Set mOptionalCalendars(i) = GetCalendar(CalendarTypes(i))
        Next i
    End If
    OptionalCalendars = mOptionalCalendars
End Property

''
' Returns the Locale ID for this culture instance.
'
' @return The Locale ID.
' @remarks A 32 bit LCID is in the format of:
'        31 - 20 19 18 17 16 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0
'        +-----+ +---------+ +---------------+ +-----------------+
'           |         |           |            Primary language ID (10 bits)
'           |         |           +----------- Sublanguage ID (6 its)
'           |         +----------------------- Sort ID (4 bits)
'           +--------------------------------- Reserved (12 bits)
'
Public Property Get LCID() As Long
    VerifyLoaded
    LCID = mProps.LCID
End Property

''
' Returns the name of the culture.
'
' @return The name of the culture.
' @remarks The culture name follows the RFC 1766 standard in that there are
' 2 parts. A 2 letter culture and 2 letter region:
' en-US = English (United States)
'
Public Property Get Name() As String
    VerifyLoaded
    Name = mProps.Name
End Property

''
' Returns the english translation for the culture name.
'
' @return The name in the english language.
'
Public Property Get EnglishName() As String
    VerifyLoaded
    EnglishName = mProps.EnglishName
End Property

''
' Returns a displayable name for the culture name.
'
' @return The name as it should be displayed.
'
Public Property Get DisplayName() As String
    VerifyLoaded
    DisplayName = mProps.DisplayName
End Property

''
' Returns if this culture is neutral.
'
' @return Indication of the culture's neutrality.
' @remarks A neutral culture is a culture that has no region specified.
' en is a neutral culture, where as en-US is not.
'
Public Property Get IsNeutralCulture() As Boolean
    VerifyLoaded
    If mProps.LCID <> INVARIANT_LCID Then
        IsNeutralCulture = CultureTable.IsNeutral(mProps.LCID)
    End If
End Property

''
' Returns if this instance is ReadOnly.
'
' @return The readonly status of this instance.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

Friend Sub MarkReadOnly()
    mProps.IsReadOnly = True
    
    If Not mDateTimeFormat Is Nothing Then
        mDateTimeFormat.MarkReadOnly
    End If
    
    If Not mNumberFormat Is Nothing Then
        mNumberFormat.MarkReadOnly
    End If
End Sub

''
' Returns the native name for this culture instance.
'
' @return The name using native characters to spell the culture name.
'
Public Property Get NativeName() As String
    VerifyLoaded
    NativeName = mProps.NativeName
End Property

''
' Returns the 3 letter ISO 639-2 standard of the culture name.
'
' @return The name of the culture in three letters.
'
Public Property Get ThreeLetterISOLanguageName() As String
    VerifyLoaded
    ThreeLetterISOLanguageName = mProps.ThreeLetterISOLanguageName
End Property

''
' Returns the 3 letter Windows name of the culture name.
'
' @return The name of the culture found in windows.
'
Public Property Get ThreeLetterWindowsLanguageName() As String
    VerifyLoaded
    ThreeLetterWindowsLanguageName = mProps.ThreeLetterWindowLanguageName
End Property

''
' Returns the 2 letter  ISO 639-1 standard of the culture name.
'
' @return The name of the culture in 2 letters.
'
Public Property Get TwoLetterISOLanguageName() As String
    VerifyLoaded
    TwoLetterISOLanguageName = mProps.TwoLetterISOLanguageName
End Property

''
' Returns whether the culture instance is using the user settings.
'
' @return Indication as to if the culture is using the user setting
' defined in the control panel.
'
Public Property Get UseUserOverride() As Boolean
    UseUserOverride = mProps.UseUserOverride
End Property

''
' Returns a clone of this culture instance.
'
' @return An instance of this culture. ReadOnly status is retained.
'
Public Function Clone() As CultureInfo
    Dim DTInfo As DateTimeFormatInfo
    If Not mDateTimeFormat Is Nothing Then
        Set DTInfo = mDateTimeFormat.Clone
    End If
    
    Dim NumInfo As NumberFormatInfo
    If Not mNumberFormat Is Nothing Then
        Set NumInfo = mNumberFormat.Clone
    End If
        
    Dim Bag As New PropertyBag
    Bag.WriteProperty "Calendar", mCalendar
    
    Dim CalInfo As Calendar
    Set CalInfo = Bag.ReadProperty("Calendar", Nothing)
    
    Set Clone = New CultureInfo
    Clone.InitClone mProps, DTInfo, NumInfo, CalInfo
End Function

''
' Returns a format provider of the specified type.
'
' @param FormatType The type of format provider that is requested.
' @return The format provider, or Nothing if a provider was not found.
'
Public Function GetFormat(ByRef FormatType As String) As Object
    If CorString.Equals(FormatType, vbNumberFormatInfo, OrdinalIgnoreCase) Then
        Set GetFormat = NumberFormat
    ElseIf CorString.Equals(FormatType, vbDateTimeFormatInfo, OrdinalIgnoreCase) Then
        Set GetFormat = DateTimeFormat
    End If
End Function

''
' Returns the NumberFormatInfo associated with this culture.
'
' @return The NumberFormatInfo for this culture that can be used
' to format numbers specific to this culture.
'
Public Property Get NumberFormat() As NumberFormatInfo
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
    
    If mNumberFormat Is Nothing Then
        Set mNumberFormat = New NumberFormatInfo
        mNumberFormat.InitCultureInfo Me
    End If
    Set NumberFormat = mNumberFormat
End Property

''
' Sets the NumberFormatInfo for the specific culture object.
'
' @param Value The NumberFormatInfo to associate with this object instance.
'
Public Property Set NumberFormat(ByVal Value As NumberFormatInfo)
    VerifyWritable
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
        
    Set mNumberFormat = Value
End Property

''
' Returns the DateTimeFormatInfo associated with this culture.
'
' @return The DateTimeFormatInfo for this culture that can be used
' to format dates and times specific to this culture.
'
Public Property Get DateTimeFormat() As DateTimeFormatInfo
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
        
    If mDateTimeFormat Is Nothing Then
        Set mDateTimeFormat = New DateTimeFormatInfo
        mDateTimeFormat.InitCultureInfo Me
    End If
    Set DateTimeFormat = mDateTimeFormat
End Property

''
' Sets the DateTimeFormatInfo for the specific culture object.
'
' @param Value The DateTimeFormatInfo object to associate with this object instance.
'
Public Property Set DateTimeFormat(ByVal Value As DateTimeFormatInfo)
    VerifyWritable
    If Value Is Nothing Then _
        Error.ArgumentNull "Value"
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
    
    Set mDateTimeFormat = Value
End Property

''
' Returns the parent culture for this culture.
'
' @return The parent culture for this culture. If this culture is
' invariant, then invariant is returned.
'
Public Property Get Parent() As CultureInfo
    Set Parent = Cor.NewCultureInfo(mProps.ParentLCID)
End Property

''
' Returns a string representation of this object instance.
'
' @return String representing this instance.
Public Function ToString() As String
    ToString = Me.Name
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
Public Function Equals(ByRef Value As Variant) As Boolean
    Dim Other As CultureInfo
    
    On Error GoTo Catch
    Set Other = Value
    Equals = Other.Name = mProps.Name

Catch:
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = Object.GetHashCode(mProps.Name)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef NameOrLCID As Variant, ByVal UseUserOverride As Boolean)
    Dim ID As Long
    
    Select Case VarType(NameOrLCID)
        Case vbLong, vbInteger, vbByte
            ID = NameOrLCID
        Case vbString
            If Len(NameOrLCID) > 0 Then
                If CultureTable.IsLoaded Then
                    ID = CultureTable.GetCultureID(CStr(NameOrLCID))
                End If
            Else
                ID = INVARIANT_LCID
            End If
        Case Else
            Throw Cor.NewArgumentException("Invalid Culture Identifier.", "NameOrLCID")
    End Select
    
    If ID <> INVARIANT_LCID Then
        mProps.UseUserOverride = UseUserOverride
    Else
        mProps.UseUserOverride = False
    End If
    
    If ID < 0 Then _
        Error.ArgumentOutOfRange "LCID", ArgumentOutOfRange_NeedNonNegNum
    Load ID
End Sub

Friend Sub InitClone(ByRef Props As PropsType, ByVal DTInfo As DateTimeFormatInfo, ByVal NumInfo As NumberFormatInfo, ByVal CalInfo As Calendar)
    mProps = Props
    Set mDateTimeFormat = DTInfo
    Set mNumberFormat = NumInfo
    Set mCalendar = CalInfo
    mLoaded = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyWritable()
    If mProps.IsReadOnly Then Throw Cor.NewInvalidOperationException("Culture is Read-Only.")
End Sub

Private Sub VerifyLoaded()
    If Not mLoaded Then
        Load INVARIANT_LCID
    End If
End Sub

Private Sub Load(ByVal LCID As Long)
    If CultureTable.IsLoaded Then
        With mProps
            .LCID = LCID
            .ParentLCID = CultureTable.GetNumber(LCID, IPARENTLCID)
            .Name = CultureTable.GetString(LCID, SNAME)
            .EnglishName = CultureTable.GetString(LCID, SENGLISHNAME)
            .DisplayName = CultureTable.GetString(LCID, SDISPLAYNAME)
            .NativeName = CultureTable.GetString(LCID, SNATIVENAME)
            .ThreeLetterISOLanguageName = CultureTable.GetString(LCID, STHREELETTERISOLANGUAGENAME)
            .ThreeLetterWindowLanguageName = CultureTable.GetString(LCID, STHREELETTERWINDOWSLANGUAGENAME)
            .TwoLetterISOLanguageName = CultureTable.GetString(LCID, STWOLETTERISOLANGUAGENAME)
        End With
    Else
        LoadDefault
    End If
    
    mLoaded = True
End Sub

Private Sub LoadDefault()
    With mProps
        .LCID = INVARIANT_LCID
        .ParentLCID = INVARIANT_LCID
        .Name = ""
        .EnglishName = "Invariant Language (Invariant Country)"
        .DisplayName = "Invariant Language (Invariant Country)"
        .NativeName = "Invariant Language (Invariant Country)"
        .ThreeLetterISOLanguageName = "IVL"
        .ThreeLetterWindowLanguageName = "IVL"
        .TwoLetterISOLanguageName = "iv"
    End With
End Sub

Private Function GetCalendar(ByVal CalendarType As Long) As Calendar
    Select Case CalendarType
        Case CAL_GREGORIAN, _
             CAL_GREGORIAN_US, _
             CAL_GREGORIAN_ME_FRENCH, _
             CAL_GREGORIAN_ARABIC, _
             CAL_GREGORIAN_XLIT_ENGLISH, _
             CAL_GREGORIAN_XLIT_FRENCH
            
            Dim g As New GregorianCalendar
            g.CalendarType = CalendarType
            Set GetCalendar = g
            
        Case CAL_JAPAN:     Set GetCalendar = New JapaneseCalendar
        Case CAL_TAIWAN:    Set GetCalendar = New TaiwanCalendar
        Case CAL_KOREA:     Set GetCalendar = New KoreanCalendar
        Case CAL_HIJRI:     Set GetCalendar = New HijriCalendar
        Case CAL_THAI:      Set GetCalendar = New ThaiBuddhistCalendar
        Case CAL_HEBREW:    Set GetCalendar = New HebrewCalendar
        Case CAL_JULIAN:    Set GetCalendar = New JulianCalendar
        Case Else:          Set GetCalendar = New GregorianCalendar
    End Select
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mProps.LCID = .ReadProperty(PROP_LCID)
        mProps.EnglishName = .ReadProperty(PROP_ENGLISHNAME)
        mProps.DisplayName = .ReadProperty(PROP_DISPLAYNAME)
        mProps.NativeName = .ReadProperty(PROP_NATIVENAME)
        mProps.Name = .ReadProperty(PROP_NAME)
        mProps.ParentLCID = .ReadProperty(PROP_PARENTLCID)
        mProps.ThreeLetterISOLanguageName = .ReadProperty(PROP_THREELETTERISOLANGUAGENAME)
        mProps.ThreeLetterWindowLanguageName = .ReadProperty(PROP_THREELETTERWINDOWSLANGUAGENAME)
        mProps.TwoLetterISOLanguageName = .ReadProperty(PROP_TWOLETTERISOLANGUAGENAME)
        mProps.UseUserOverride = .ReadProperty(PROP_USEUSEROVERRIDE)
        mProps.IsReadOnly = .ReadProperty(PROP_ISREADONLY)
        Set mDateTimeFormat = .ReadProperty(PROP_DATETIMEFORMAT)
        Set mNumberFormat = .ReadProperty(PROP_NUMBERFORMAT)
        Set mCalendar = .ReadProperty(PROP_CALENDAR)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PROP_LCID, mProps.LCID
        .WriteProperty PROP_PARENTLCID, mProps.ParentLCID
        .WriteProperty PROP_NAME, mProps.Name
        .WriteProperty PROP_ENGLISHNAME, mProps.EnglishName
        .WriteProperty PROP_DISPLAYNAME, mProps.DisplayName
        .WriteProperty PROP_NATIVENAME, mProps.NativeName
        .WriteProperty PROP_THREELETTERISOLANGUAGENAME, mProps.ThreeLetterISOLanguageName
        .WriteProperty PROP_THREELETTERWINDOWSLANGUAGENAME, mProps.ThreeLetterWindowLanguageName
        .WriteProperty PROP_TWOLETTERISOLANGUAGENAME, mProps.TwoLetterISOLanguageName
        .WriteProperty PROP_USEUSEROVERRIDE, mProps.UseUserOverride
        .WriteProperty PROP_ISREADONLY, mProps.IsReadOnly
        .WriteProperty PROP_DATETIMEFORMAT, mDateTimeFormat
        .WriteProperty PROP_NUMBERFORMAT, mNumberFormat
        .WriteProperty PROP_CALENDAR, mCalendar
    End With
End Sub


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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function

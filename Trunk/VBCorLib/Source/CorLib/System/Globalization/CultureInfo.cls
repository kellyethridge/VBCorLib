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
'    Module: CultureInfo
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
    LCID As Long
    ParentLCID As Long
    Name As String
    EnglishName As String
    DisplayName As String
    NativeName As String
    ThreeLetterISOLanguageName As String
    ThreeLetterWindowLanguageName As String
    TwoLetterISOLanguageName As String
    UseUserOverride As Boolean
    
    IsReadOnly As Boolean
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
    If mCalendar Is Nothing Then Set mCalendar = GetCalendar(CultureTable.GetNumber(mProps.LCID, ICALENDARTYPE))
    Set Calendar = mCalendar
End Property

''
' Returns a list of optional calendars supported by the culture.
'
' @return A list of calendars associated with the current culture.
'
Public Property Get OptionalCalendars() As Calendar()
    If cArray.IsNull(mOptionalCalendars) Then
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
    Call VerifyLoaded
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
    Call VerifyLoaded
    Name = mProps.Name
End Property

''
' Returns the english translation for the culture name.
'
' @return The name in the english language.
'
Public Property Get EnglishName() As String
    Call VerifyLoaded
    EnglishName = mProps.EnglishName
End Property

''
' Returns a displayable name for the culture name.
'
' @return The name as it should be displayed.
'
Public Property Get DisplayName() As String
    Call VerifyLoaded
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
    Call VerifyLoaded
    If mProps.LCID = INVARIANT_LCID Then Exit Property
    IsNeutralCulture = CultureTable.IsNeutral(mProps.LCID)
End Property

''
' Returns if this instance is ReadOnly.
'
' @return The readonly status of this instance.
'
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mProps.IsReadOnly
End Property

''
' Returns the native name for this culture instance.
'
' @return The name using native characters to spell the culture name.
'
Public Property Get NativeName() As String
    Call VerifyLoaded
    NativeName = mProps.NativeName
End Property

''
' Returns the 3 letter ISO 639-2 standard of the culture name.
'
' @return The name of the culture in three letters.
'
Public Property Get ThreeLetterISOLanguageName() As String
    Call VerifyLoaded
    ThreeLetterISOLanguageName = mProps.ThreeLetterISOLanguageName
End Property

''
' Returns the 3 letter Windows name of the culture name.
'
' @return The name of the culture found in windows.
'
Public Property Get ThreeLetterWindowsLanguageName() As String
    Call VerifyLoaded
    ThreeLetterWindowsLanguageName = mProps.ThreeLetterWindowLanguageName
End Property

''
' Returns the 2 letter  ISO 639-1 standard of the culture name.
'
' @return The name of the culture in 2 letters.
'
Public Property Get TwoLetterISOLanguageName() As String
    Call VerifyLoaded
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
    Dim dtInfo As DateTimeFormatInfo
    If Not mDateTimeFormat Is Nothing Then Set dtInfo = mDateTimeFormat.Clone
    
    Dim nInfo As NumberFormatInfo
    If Not mNumberFormat Is Nothing Then Set nInfo = mNumberFormat.Clone
        
    Dim pb As New PropertyBag
    Call pb.WriteProperty("Calendar", mCalendar)
    
    Dim calInfo As Calendar
    Set calInfo = pb.ReadProperty("Calendar", Nothing)
    
    Set Clone = New CultureInfo
    Call Clone.CloneHelper(mProps, dtInfo, nInfo, calInfo)
End Function

''
' Returns a format provider of the specified type.
'
' @param FormatType The type of format provider that is requested.
' @return The format provider, or Nothing if a provider was not found.
'
Public Function GetFormat(ByRef FormatType As String) As Object
    If StringUtils.EqualText(FormatType, "numberformatinfo") Then
        Set GetFormat = NumberFormat
    ElseIf StringUtils.EqualText(FormatType, "datetimeformatinfo") Then
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
        Call mNumberFormat.Load(mProps.LCID, mProps.UseUserOverride)
        mNumberFormat.IsReadOnly = mProps.IsReadOnly
    End If
    Set NumberFormat = mNumberFormat
End Property

''
' Sets the NumberFormatInfo for the specific culture object.
'
' @param RHS The NumberFormatInfo to associate with this object instance.
'
Public Property Set NumberFormat(ByVal RHS As NumberFormatInfo)
    Call VerifyWritable
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("NumberFormat")
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
        
    Set mNumberFormat = RHS
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
        Call mDateTimeFormat.Load(mProps.LCID, mProps.UseUserOverride, Me.Calendar)
        mDateTimeFormat.IsReadOnly = mProps.IsReadOnly
    End If
    Set DateTimeFormat = mDateTimeFormat
End Property

''
' Sets the DateTimeFormatInfo for the specific culture object.
'
' @param RHS The DateTimeFormatInfo object to associate with this object instance.
'
Public Property Set DateTimeFormat(ByVal RHS As DateTimeFormatInfo)
    Call VerifyWritable
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("DateTimeFormat")
    If IsNeutralCulture Then _
        Throw Cor.NewNotSupportedException("Neutral cultures cannot provide formatting.")
    
    Set mDateTimeFormat = RHS
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
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Function
        
        Dim c As CultureInfo
        If TypeOf Value Is CultureInfo Then
            Set c = Value
            Equals = (c.LCID = mProps.LCID)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
Public Function GetHashCode() As Long
    GetHashCode = Me.LCID
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef NameOrLCID As Variant, ByVal UseUserOverride As Boolean)
    Dim ID As Long
    
    Select Case VarType(NameOrLCID)
        Case vbLong, vbInteger, vbByte
            ID = NameOrLCID
        Case vbString
            If CultureTable.IsLoaded Then ID = CultureTable.GetCultureID(NameOrLCID)
        Case Else
            Throw Cor.NewArgumentException("Invalid Culture Identifier.", "NameOrLCID")
    End Select
    If UseUserOverride Then mProps.UseUserOverride = (GetUserDefaultLCID = ID)
    Call Load(ID)
End Sub

Friend Sub CloneHelper(ByRef props As PropsType, ByVal dtInfo As DateTimeFormatInfo, ByVal nInfo As NumberFormatInfo, ByVal calInfo As Calendar)
    mProps = props
    Set mDateTimeFormat = dtInfo
    Set mNumberFormat = nInfo
    Set mCalendar = calInfo
End Sub

Friend Property Let IsReadOnly(ByVal RHS As Boolean)
    mProps.IsReadOnly = RHS
    If Not mDateTimeFormat Is Nothing Then mDateTimeFormat.IsReadOnly = RHS
    If Not mNumberFormat Is Nothing Then mNumberFormat.IsReadOnly = RHS
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyWritable()
    If mProps.IsReadOnly Then Throw Cor.NewInvalidOperationException("Culture is Read-Only.")
End Sub

Private Sub VerifyLoaded()
    If Not mLoaded Then Call Load(INVARIANT_LCID)
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
        Call LoadDefault
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
'   Class Events
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
        Call .WriteProperty(PROP_LCID, mProps.LCID)
        Call .WriteProperty(PROP_PARENTLCID, mProps.ParentLCID)
        Call .WriteProperty(PROP_NAME, mProps.Name)
        Call .WriteProperty(PROP_ENGLISHNAME, mProps.EnglishName)
        Call .WriteProperty(PROP_DISPLAYNAME, mProps.DisplayName)
        Call .WriteProperty(PROP_NATIVENAME, mProps.NativeName)
        Call .WriteProperty(PROP_THREELETTERISOLANGUAGENAME, mProps.ThreeLetterISOLanguageName)
        Call .WriteProperty(PROP_THREELETTERWINDOWSLANGUAGENAME, mProps.ThreeLetterWindowLanguageName)
        Call .WriteProperty(PROP_TWOLETTERISOLANGUAGENAME, mProps.TwoLetterISOLanguageName)
        Call .WriteProperty(PROP_USEUSEROVERRIDE, mProps.UseUserOverride)
        Call .WriteProperty(PROP_ISREADONLY, mProps.IsReadOnly)
        Call .WriteProperty(PROP_DATETIMEFORMAT, mDateTimeFormat)
        Call .WriteProperty(PROP_NUMBERFORMAT, mNumberFormat)
        Call .WriteProperty(PROP_CALENDAR, mCalendar)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function

Private Function IObject_GetHashcode() As Long
    IObject_GetHashcode = GetHashCode
End Function

Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IFormatProvider Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IFormatProvider_GetFormat(ByRef FormatType As String) As Object
    Set IFormatProvider_GetFormat = GetFormat(FormatType)
End Function

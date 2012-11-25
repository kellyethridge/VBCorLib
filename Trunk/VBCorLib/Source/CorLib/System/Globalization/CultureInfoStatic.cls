VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfoStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: CultureInfoStatic
'

''
' Provides a set of static methods retrieve CultureInfo objects.
'
' @remarks This class cannot be instantiated. In order to access the methods
' use the variable name directly.
' <pre>
' Dim Cultures() As CultureInfo
' Cultures = Culture.GetCultures(AllCultures)
' </pre>
'
' @see CultureInfo
'
Option Explicit

''
' A list of culture type filters used in <b>GetCultures</b>.
'
' @param NeutralCultures 1) Refers to cultures that are associated with a language but are not specific to a country/region. The names of these cultures consist of the lowercase two-letter code derived from ISO 639-1. For example: "en" (English) is a neutral culture. The invariant culture is included in the array of cultures returned by CultureInfo.GetCultures with this value.
' @param SpecificCultures 2) Refers to cultures that are specific to a country/region. The names of these cultures follow the RFC 1766 standard in the format "<languagecode2>-<country/regioncode2>", where <languagecode2> is a lowercase two-letter code derived from ISO 639-1 and <country/regioncode2> is an uppercase two-letter code derived from ISO 3166. For example, "en-US" (English - United States) is a specific culture.
' @param InstalledWin32Cultures 4) Refers to all cultures that are installed in the Windows system. Note that not all cultures supported by the .NET Framework are installed in the Windows system.
' @param AllCultures 7) Refers to all cultures.
'
Public Enum CultureTypes
    NeutralCultures = 1
    SpecificCultures = 2
    InstalledWin32Cultures = 4
    AllCultures = 7
End Enum

Private mInvariantCulture   As CultureInfo
Private mCurrentUICulture   As CultureInfo
Private mInstalledUICulture As CultureInfo



''
' Returns a ReadOnly wrapped CultureInfo object.
'
' @param Culture The culture to create a ReadOnly wrapper for.
' @return The wrapped ReadOnly culture.
'
Public Function ReadOnly(ByVal Culture As CultureInfo) As CultureInfo
    If Culture.IsReadOnly Then
        Set ReadOnly = Culture
    Else
        Set ReadOnly = Culture.Clone
        ReadOnly.IsReadOnly = True
    End If
End Function

''
' Returns a set of CultureInfo objects.
'
' @param Types The culture types that are to be retrieved.
' @return An array of CultureInfo objects based on Types specified.
'
Public Function GetCultures(ByVal Types As CultureTypes) As CultureInfo()
    GetCultures = CultureTable.GetCultures(Types)
End Function

''
' Returns the current culture associated with this machine.
'
' @return A ReadOnly version of CultureInfo specific for this machine.
'
Public Property Get CurrentCulture() As CultureInfo
    Set CurrentCulture = Thread.CurrentThread.CurrentCulture
End Property

''
' Returns a ReadOnly version of the invariant culture.
'
' @return A ReadOnly version of the invariant culture.
'
Public Property Get InvariantCulture() As CultureInfo
    If mInvariantCulture Is Nothing Then
        Set mInvariantCulture = Cor.NewCultureInfo(INVARIANT_LCID)
        mInvariantCulture.IsReadOnly = True
    End If
    Set InvariantCulture = mInvariantCulture
End Property

''
' Returns the culture for the current thread.
'
' @return A <b>CultureInfo</b> object for the current thread.
' @remarks The culture is determined by the language ID of the
' current thread by calling <b>GetUserDefaultUILanguage</b> API.
'
Public Property Get CurrentUICulture() As CultureInfo
    If mCurrentUICulture Is Nothing Then
        Set mCurrentUICulture = Cor.NewCultureInfo(GetUserDefaultUILanguage)
        mCurrentUICulture.IsReadOnly = True
    End If
    Set CurrentUICulture = mCurrentUICulture
End Property

''
' Returns the culture for the current systems language.
'
' @return A <b>CultureInfo</b> object for the current system language.
' @remarks The current system language is detemined by calling the
' <b>GetSystemDefaultUILanguage</b> API.
'
Public Property Get InstalledUICulture() As CultureInfo
    If mInstalledUICulture Is Nothing Then
        Set mInstalledUICulture = Cor.NewCultureInfo(GetSystemDefaultUILanguage)
        mInstalledUICulture.IsReadOnly = True
    End If
    Set InstalledUICulture = mInstalledUICulture
End Property



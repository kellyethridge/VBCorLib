VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2019 Kelly Ethridge
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
' Module: OperatingSystem
'

''
' Provides information about the currently running operating system.
'
' @see Constructors
' @see ICloneable
'
'@Folder("CorLib.System")
Option Explicit
Implements IObject
Implements ICloneable


''
' The identification for the specific Windows platforms.
'
' @param Win32S Identification for legacy Windows systems that are 16bit, but allowed to access 32bit functionality.
' @param Win32Windows Identification for Windows 95, 98, ME platforms.
' @param Win32NT Identification for Windows NT4.0, 2000, XP, Server 2003, and future NT platforms.
' @param WinCE Identification for Windows CE which runs on handheld PDA's.
'
Public Enum PlatformID
    Win32S
    Win32Windows
    Win32NT
    WinCE
End Enum


Private mPlatformID     As PlatformID
Private mVersion        As Version
Private mServicePack    As String
Private mVersionString  As String


''
' Returns the service pack for the current operating system.
'
' @remarks An empty string is always returned on Windows 95, 98, and ME.
'
Public Property Get ServicePack() As String
    ServicePack = mServicePack
End Property

''
' Returns the operating system type currently running on this machine.
'
' @return The current operating system running.
'
Public Property Get Platform() As PlatformID
    Platform = mPlatformID
End Property

''
' Returns the version of the current operating system.
'
' @return The operating system version.
'
Public Property Get Version() As Version
    Set Version = mVersion
End Property

''
' Returns a clone of the OperatingSystem object.
'
' @return The clone of the OperatingSystem object.
'
Public Function Clone() As OperatingSystem
    Set Clone = Cor.NewOperatingSystem(mPlatformID, mVersion, mServicePack)
End Function

''
' Gets the concatenated string representation of the platform identifier, version, and service pack that are currently installed on the operating system.
'
' @return The string representation of the values returned by the <a href="OperatingSystem.get.Platform.html">Platform</a>, <a href="OperatingSystem.get.Version.html">Version</a>, and <a href="OperatingSystem.get.ServicePack.html">ServicePack</a> properties.
'
Public Property Get VersionString() As String
    Dim OSName As String
    
    If Len(mVersionString) = 0 Then
        Select Case mPlatformID
            Case PlatformID.Win32NT
                OSName = "Microsoft Windows NT"
            Case PlatformID.Win32S
                OSName = "Microsoft Win32S"
            Case PlatformID.Win32Windows
                If mVersion.Major = 4 And mVersion.Minor = 0 Then
                    OSName = "Microsoft Windows 95"
                Else
                    OSName = "Microsoft Windows 98"
                End If
            Case PlatformID.WinCE
                OSName = "Microsoft Windows CE"
            Case Else
                OSName = "<unknown>"
        End Select
        
        If Len(mServicePack) = 0 Then
            mVersionString = CorString.Format("{0} {1}", OSName, mVersion)
        Else
            mVersionString = CorString.Format("{0} {1} {2}", OSName, mVersion.ToString(3), mServicePack)
        End If
    End If
    
    VersionString = mVersionString
End Property

''
' Returns a string representation of the current operating system.
'
' @return String representing this instance.
'
Public Function ToString() As String
    ToString = VersionString
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    Equals = MyBase.Equals(Me, Value)
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    GetHashCode = MyBase.GetHashCode(Me)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mPlatformID = .ReadProperty("PlatformID", 0)
        Set mVersion = .ReadProperty("Version", Nothing)
        mServicePack = .ReadProperty("ServicePack", "")
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "PlateformID", mPlatformID
        .WriteProperty "Version", mVersion
        .WriteProperty "ServicePack", mServicePack
    End With
End Sub

Friend Sub Init(ByVal Platform As PlatformID, ByVal Version As Version, ByRef ServicePack As String)
    mPlatformID = Platform
    Set mVersion = Version
    mServicePack = ServicePack
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

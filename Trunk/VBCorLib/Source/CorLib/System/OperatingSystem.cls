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
'    Module: OperatingSystem
'

''
' Provides information about the currently running operating system.
'
' @see Constructors
' @see ICloneable
'
Option Explicit
Implements IObject
Implements ICloneable



''
' The identification for the specific Windows platforms.
'
' @param Win32NT Identification for Windows NT4.0, 2000, XP, Server 2003,
' and future NT platforms.
' @param Win32S Identification for legacy Windows systems that are 16bit,
' but allowed to access 32bit functionality.
' @param Win32Windows Identification for Windows 95, 98, ME platforms.
' @param WinCE Identification for Windows CE which runs on handheld PDA's.
'
Public Enum PlatformID
   Win32NT = 2
   Win32S = 0
   Win32Windows = 1
   WinCE = 3
End Enum


Private mPlatformID     As PlatformID
Private mVersion        As Version
Private mLoaded         As Boolean
Private mServicePack    As String


''
' Returns the service pack for the current operating system.
'
' @remarks An empty string is always returned on Windows 95, 98, and ME.
'
Public Property Get ServicePack() As String
    Call VerifyLoaded
    ServicePack = mServicePack
End Property

''
' Returns the operating system type currently running on this machine.
'
' @return The current operating system running.
'
Public Property Get Platform() As PlatformID
    Call VerifyLoaded
    Platform = mPlatformID
End Property

''
' Returns the version of the current operating system.
'
' @return The operating system version.
'
Public Property Get Version() As Version
    Call VerifyLoaded
    Set Version = mVersion
End Property

''
' Returns a clone of the OperatingSystem object.
'
' @return The clone of the OperatingSystem object.
'
Public Function Clone() As OperatingSystem
    Call VerifyLoaded
    Set Clone = Cor.NewOperatingSystem(mPlatformID, mVersion.Clone, mServicePack)
End Function

''
' Returns a string representation of the current operating system.
'
' @return String representing this instance.
'
Public Function ToString() As String
    Call VerifyLoaded
    
    Dim Ret As String
    Select Case mPlatformID
        Case Win32NT
            If mVersion.Major = 5 Then
                Select Case mVersion.Minor
                    Case 0:     Ret = "Microsoft Windows 2000"
                    Case 1:     Ret = "Microsoft Windows XP"
                    Case 2:     Ret = "Microsoft Windows Server 2003 family"
                    Case Else:  Ret = "Microsoft Windows NT"
                End Select
            Else
                Ret = "Microsoft Windows NT"
            End If
        Case Win32S
            Ret = "Microsoft Win32s"
        Case Win32Windows
            Select Case mVersion.Minor
                Case 0:     Ret = "Microsoft Windows 95"
                Case 10:    Ret = "Micorosft Windows 98"
                Case 90:    Ret = "Micorsoft Windows ME"
                Case Else:  Ret = "Microsoft Windows32"
            End Select
        Case WinCE
            Ret = "Microsoft WinCE"
    End Select
    
    If Len(mServicePack) > 0 Then
        ToString = Ret & " " & mVersion.ToString(3) & " " & mServicePack
    Else
        ToString = Ret & " " & mVersion.ToString
    End If
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
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Platform As PlatformID, ByVal Version As Version, ByRef ServicePack As String)
    mPlatformID = Platform
    Set mVersion = Version
    mServicePack = ServicePack
    mLoaded = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifyLoaded()
    If mLoaded Then Exit Sub
    
    Dim Info As OSVERSIONINFOA
    Info.dwOSVersionInfoSize = Len(Info)
    If GetVersionExA(Info) = BOOL_FALSE Then _
        Throw Cor.NewInvalidOperationException("Could not load Operating System information.")
    
    ' We are on an NT platform so we can get extra information
    If Info.dwPlatformId = PlatformID.Win32NT Then
        Dim infoex As OSVERSIONINFOEXA
        infoex.dwOSVersionInfoSize = Len(infoex)
        If GetVersionExA(infoex) = BOOL_FALSE Then _
            Throw Cor.NewInvalidOperationException("Could not load Operating System information.")
        
        Dim ServicePack As String
        ServicePack = StrConv(infoex.szCSDVersion, vbUnicode)
        ServicePack = SysAllocString(StrPtr(ServicePack))
    End If
    Call Init(Info.dwPlatformId, Cor.NewVersion(Info.dwMajorVersion, Info.dwMinorVersion, Info.dwBuildNumber, Helper.ShiftLeft(infoex.wServicePackMajor, 16) Or infoex.wServicePackMinor), ServicePack)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mPlatformID = .ReadProperty("PlatformID", 0)
        Set mVersion = .ReadProperty("Version", Nothing)
        mServicePack = .ReadProperty("ServicePack", "")
    End With
    mLoaded = True
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("PlateformID", mPlatformID)
        Call .WriteProperty("Version", mVersion)
        Call .WriteProperty("ServicePack", mServicePack)
    End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IObject Interface
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
'   ICloneable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
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
'    Module: Version
'

''
' Represents a version number.
'
' @see Constructors
' @see VersionStatic
'
Option Explicit
Implements IObject
Implements ICloneable
Implements IComparable

Private Const PROP_MAJOR            As String = "Major"
Private Const PROP_MINOR            As String = "Minor"
Private Const PROP_BUILD            As String = "Build"
Private Const PROP_REVISION         As String = "Revision"

Private Const UNUSED                As Long = -1
Private Const DEF_MAJOR             As Long = 0
Private Const DEF_MINOR             As Long = 0
Private Const DEF_BUILD             As Long = UNUSED
Private Const DEF_REVISION          As Long = UNUSED


Private mMajor      As Long
Private mMinor      As Long
Private mBuild      As Long
Private mRevision   As Long


''
' Returns the major version number of this instance.
'
' @return The major version number. This is a non-negative number.
'
Public Property Get Major() As Long
    Major = mMajor
End Property

''
' Returns the minor version number of this instance.
'
' @return The minor version number. This is a non-negative number.
'
Public Property Get Minor() As Long
    Minor = mMinor
End Property

''
' Returns the build version number of this instance.
'
' @return The build version number.
' @remarks The build version is an optional tracking number. If the number
' was not used when this instance was created, then -1 is returned.
'
Public Property Get Build() As Long
    Build = mBuild
End Property

''
' Returns the revision version number of this instance.
'
' @return The revision version number.
' @remarks The build version is an optional tracking number. If the number
' was not used when this instance was created, then -1 is returned.
'
Public Property Get Revision() As Long
    Revision = mRevision
End Property

''
' Returns a minor revision value.
'
' @return Minor revision number.
' @remarks The minor revision is the lower 16-bits of the
' standard revision value.
'
Public Property Get MinorRevision() As Long
    MinorRevision = mRevision And &HFFFF&
End Property

''
' Returns the major revision value.
'
' @return Major revision number.
' @remarks The major revision number is the upper 16-bits
' of the standard revision value.
'
Public Property Get MajorRevision() As Long
    MajorRevision = Helper.ShiftRight(mRevision, 16)
End Property

''
' Returns a cloned object of this instance.
'
' @return The cloned version of this instanced.
'
Public Function Clone() As Version
    Set Clone = New Version
    Call Clone.CloneHelper(mMajor, mMinor, mBuild, mRevision)
End Function

''
' Compares this instance to a passed in value.
'
' @param value The value to compare this instance against.
' @return Value indicating the relation of this instance to the passed in value.
' @remarks A negative number indicates this instance is less than the value.
' A positive number indicates this instance is greater than the value.
' Zero indicates this instance is equal to the value.
'
Public Function CompareTo(ByRef Value As Variant) As Long
    Select Case VarType(Value)
        Case vbObject
            If Value Is Nothing Then: CompareTo = 1: Exit Function
            If Not TypeOf Value Is Version Then Throw Cor.NewArgumentException("Value must of type Version.", "value")

            Dim v As Version
            Set v = Value
            If Not EqualFields(mMajor, v.Major, CompareTo) Then Exit Function
            If Not EqualFields(mMinor, v.Minor, CompareTo) Then Exit Function
            If Not EqualFields(mBuild, v.Build, CompareTo) Then Exit Function
            If Not EqualFields(mRevision, v.Revision, CompareTo) Then Exit Function

        Case vbNull, vbEmpty, vbError
            CompareTo = 1
        Case Else
            Throw Cor.NewArgumentException("Value must of type Version.", "value")
    End Select
End Function

''
' Returns a string representation of this object instance.
'
' @param vFieldCount The number of fields to be included in the result string.
' @return String representing this instance in the form of Major.Minor[.Build[.Revision]]
'
Public Function ToString(Optional ByRef vFieldCount As Variant) As String
    Dim Ret As String
    Dim FieldCount As Long
    Dim maxfield As Long
    
    If IsMissing(vFieldCount) Then
        FieldCount = -1
    Else
        FieldCount = vFieldCount
        maxfield = GetMaxFieldCount
        If FieldCount < 0 Or FieldCount > maxfield Then _
            Throw Cor.NewArgumentOutOfRangeException("FieldCount", FieldCount, Environment.GetResourceString(ArgumentOutOfRange_VersionFieldCount, maxfield))
    End If
    
    If FieldCount <> 0 Then Ret = mMajor
    If FieldCount < 0 Or FieldCount > 1 Then Ret = Ret & "." & mMinor
    If mBuild >= 0 And (FieldCount < 0 Or FieldCount > 2) Then
        Ret = Ret & "." & mBuild
        If mRevision >= 0 And (FieldCount < 0 Or FieldCount > 3) Then Ret = Ret & "." & mRevision
    End If
    ToString = Ret
End Function

''
' Returns a boolean indicating if the value and this object
' instance are the same instance.
'
' @param value The value to compare equality to.
' @return Boolean indicating equality.
'
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Function
        
        If TypeOf Value Is Version Then
            Equals = (CompareTo(Value) = 0)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    ' modified for Lemmy
    Dim Ret As Long
    Ret = Helper.ShiftLeft(mMajor, 24)
    Ret = Ret Or Helper.ShiftLeft(mMinor And &HFF, 16)
    Ret = Ret Or Helper.ShiftLeft(mBuild And &HFF, 8)
    GetHashCode = Ret Or (mRevision And &HFF&)
End Function

Public Function EqualTo(ByRef Value As Version) As Boolean
    If Value Is Nothing Then Exit Function
    EqualTo = (CompareTo(Value) = 0)
End Function

Public Function LessThan(ByRef Value As Version) As Boolean
    If Value Is Nothing Then _
        Throw Cor.NewArgumentNullException("Value")
    
    LessThan = (CompareTo(Value) < 0)
End Function

Public Function LessThanOrEqualTo(ByRef Value As Version) As Boolean
    If Value Is Nothing Then _
        Throw Cor.NewArgumentNullException("Value")
    
    LessThanOrEqualTo = (CompareTo(Value) <= 0)
End Function

Public Function GreaterThan(ByRef Value As Version) As Boolean
    GreaterThan = Not LessThanOrEqualTo(Value)
End Function

Public Function GreaterThanOrEqualTo(ByRef Value As Version) As Boolean
    GreaterThanOrEqualTo = Not LessThan(Value)
End Function

Public Function NotEqualTo(ByRef Value As Version) As Boolean
    NotEqualTo = Not EqualTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Major As Long, ByVal Minor As Long, ByRef vBuild As Variant, ByRef vRevision As Variant)
    If Major < 0 Or Minor < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(IIf(Major < 0, "Major", "Minor"), Message:="Version's parameters must be equal to or greater than zero.")
    
    If Not IsMissing(vBuild) Then
        mBuild = vBuild
        If mBuild < 0 Then Throw Cor.NewArgumentOutOfRangeException("Build", Message:="Version's parameters must be equal to or greater than zero.")
    End If
    If Not IsMissing(vRevision) Then
        If mBuild < 0 Then Throw Cor.NewArgumentException("Cannot include a revision without a build value.", "Revision")
        mRevision = vRevision
        If mRevision < 0 Then Throw Cor.NewArgumentOutOfRangeException("Revision", Message:="Version's parameters must be equal to or greater than zero.")
    End If
    
    mMajor = Major
    mMinor = Minor
End Sub

Friend Sub CloneHelper(ByVal Major As Long, ByVal Minor As Long, ByVal Build As Long, ByVal Revision As Long)
    mMajor = Major
    mMinor = Minor
    mBuild = Build
    mRevision = Revision
End Sub

Friend Sub Parse(ByRef s As String)
    Dim Parts() As String
    Dim ub As Long
    
    Parts = Split(s, ".")
    ub = UBound(Parts)
    If ub < 1 Or ub > 3 Then _
        Throw Cor.NewArgumentException("String must contain 2 to 4 fields.")
    mMajor = Parts(0)
    mMinor = Parts(1)
    If mMajor < 0 Or mMinor < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException(IIf(mMajor < 0, "Major", "Minor"), Message:="Version's parameters must be equal to or greater than zero.")
    
    If ub > 1 Then
        mBuild = Parts(2)
        If mBuild < 0 Then Throw Cor.NewArgumentOutOfRangeException("Build", Message:="Version's parameters must be equal to or greater than zero.")
        If ub > 2 Then
            mRevision = Parts(3)
            If mRevision < 0 Then Throw Cor.NewArgumentOutOfRangeException("Revisiont", Message:="Version's parameters must be equal to or greater than zero.")
        End If
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMaxFieldCount() As Long
    If mBuild < 0 Then
        GetMaxFieldCount = 2
    ElseIf mRevision < 0 Then
        GetMaxFieldCount = 3
    Else
        GetMaxFieldCount = 4
    End If
End Function

Private Function EqualFields(ByVal MyField As Long, ByVal TheirField As Long, ByRef RetVal As Long) As Boolean
    If MyField < TheirField Then
        RetVal = -1
    ElseIf MyField > TheirField Then
        RetVal = 1
    Else
        RetVal = 0
        EqualFields = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mBuild = UNUSED
    mRevision = UNUSED
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMajor = .ReadProperty(PROP_MAJOR, DEF_MAJOR)
        mMinor = .ReadProperty(PROP_MINOR, DEF_MINOR)
        mBuild = .ReadProperty(PROP_BUILD, DEF_BUILD)
        mRevision = .ReadProperty(PROP_REVISION, DEF_REVISION)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROP_MAJOR, mMajor)
        Call .WriteProperty(PROP_MINOR, mMinor)
        Call .WriteProperty(PROP_BUILD, mBuild)
        Call .WriteProperty(PROP_REVISION, mRevision)
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
'   IComparable Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function

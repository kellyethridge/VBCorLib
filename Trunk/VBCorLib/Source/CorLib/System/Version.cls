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
' Module: Version
'

''
' Represents a version number in the format Major.Minor[.Build[.Revision]].
'
' @see Constructors
' @see VersionStatic
'
Option Explicit
Implements IObject
Implements ICloneable
Implements IComparable

Private Const PropMajor         As String = "Major"
Private Const PropMinor         As String = "Minor"
Private Const PropBuild         As String = "Build"
Private Const PropRevision      As String = "Revision"

Private Const UnsetComponent    As Long = -1
Private Const DefaultMajor      As Long = 0
Private Const DefaultMinor      As Long = 0
Private Const DefaultBuild      As Long = UnsetComponent
Private Const DefaultRevision   As Long = UnsetComponent

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
    Clone.InitClone mMajor, mMinor, mBuild, mRevision
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
    If Not IsObject(Value) Then _
        Error.Argument Argument_VersionRequired, "Value"
        
    If Value Is Nothing Then
        CompareTo = 1
        Exit Function
    End If
    
    If Not TypeOf Value Is Version Then _
        Error.Argument Argument_VersionRequired, "Value"

    Dim Other As Version
    Set Other = Value
    
    If Not EqualFields(mMajor, Other.Major, CompareTo) Then Exit Function
    If Not EqualFields(mMinor, Other.Minor, CompareTo) Then Exit Function
    If Not EqualFields(mBuild, Other.Build, CompareTo) Then Exit Function
    If Not EqualFields(mRevision, Other.Revision, CompareTo) Then Exit Function
End Function

''
' Returns a string representation of this object instance.
'
' @param FieldCount The number of fields to be included in the result string.
' @return String representing this instance in the form of Major.Minor[.Build[.Revision]]
'
Public Function ToString(Optional ByRef FieldCount As Variant) As String
    Dim Count       As Long
    Dim MaxCount    As Long
    
    MaxCount = GetMaxFieldCount
    Count = CLngOrDefault(FieldCount, MaxCount)
    
    If Count < 0 Or Count > MaxCount Then _
        Throw Cor.NewArgumentException(Environment.GetResourceString(ArgumentOutOfRange_Bounds_Lower_Upper, 0, MaxCount), "FieldCount")
    
    If Count >= 1 Then
        ToString = mMajor
    End If
    
    If Count >= 2 Then
        ToString = ToString & "." & mMinor
    End If
    
    If Count >= 3 Then
        ToString = ToString & "." & mBuild
    End If
    
    If Count = 4 Then
        ToString = ToString & "." & mRevision
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
    If IsObject(Value) Then
        If Value Is Nothing Then
            Exit Function
        End If
        
        If TypeOf Value Is Version Then
            Equals = EqualTo(Value)
        End If
    End If
End Function

''
' Returns a pseudo-unique number identifying this instance.
'
' @return Pseudo-unique number identifying this instance.
'
Public Function GetHashCode() As Long
    Dim Result As Long
    Result = Helper.ShiftLeft(mMajor, 24)
    Result = Result Or Helper.ShiftLeft(mMinor And &HFF, 16)
    Result = Result Or Helper.ShiftLeft(mBuild And &HFF, 8)
    GetHashCode = Result Or (mRevision And &HFF&)
End Function

Public Function EqualTo(ByVal Value As Version) As Boolean
    EqualTo = (CompareTo(Value) = 0)
End Function

Public Function LessThan(ByVal Value As Version) As Boolean
    LessThan = (CompareTo(Value) < 0)
End Function

Public Function LessThanOrEqualTo(ByVal Value As Version) As Boolean
    LessThanOrEqualTo = (CompareTo(Value) <= 0)
End Function

Public Function GreaterThan(ByVal Value As Version) As Boolean
    GreaterThan = Not LessThanOrEqualTo(Value)
End Function

Public Function GreaterThanOrEqualTo(ByVal Value As Version) As Boolean
    GreaterThanOrEqualTo = Not LessThan(Value)
End Function

Public Function NotEqualTo(ByVal Value As Version) As Boolean
    NotEqualTo = Not EqualTo(Value)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Constructors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_InitProperties()
    mBuild = DefaultBuild
    mRevision = DefaultRevision
End Sub

Friend Sub Init(ByVal Major As Long, ByVal Minor As Long, ByRef Build As Variant, ByRef Revision As Variant)
    If Major < 0 Then _
        Error.ArgumentOutOfRange "Major", ArgumentOutOfRange_NeedNonNegNum
    If Minor < 0 Then _
        Error.ArgumentOutOfRange "Minor", ArgumentOutOfRange_NeedNonNegNum
    
    mMajor = Major
    mMinor = Minor
    
    If Not IsMissing(Build) Then
        mBuild = Build
        
        If mBuild < 0 Then
            Error.ArgumentOutOfRange "Build", ArgumentOutOfRange_NeedNonNegNum
        End If
    End If
    
    If Not IsMissing(Revision) Then
        mRevision = Revision
        
        If mBuild = UnsetComponent Then _
            Error.Argument Argument_ParamRequired, "Build"
        
        If mRevision < 0 Then
            Error.ArgumentOutOfRange "Revision", ArgumentOutOfRange_NeedNonNegNum
        End If
    End If
End Sub

Friend Sub InitClone(ByVal Major As Long, ByVal Minor As Long, ByVal Build As Long, ByVal Revision As Long)
    mMajor = Major
    mMinor = Minor
    mBuild = Build
    mRevision = Revision
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMaxFieldCount() As Long
    If mBuild = UnsetComponent Then
        GetMaxFieldCount = 2
    ElseIf mRevision = UnsetComponent Then
        GetMaxFieldCount = 3
    Else
        GetMaxFieldCount = 4
    End If
End Function

Private Function EqualFields(ByVal MyField As Long, ByVal OtherField As Long, ByRef Result As Long) As Boolean
    If MyField < OtherField Then
        Result = -1
    ElseIf MyField > OtherField Then
        Result = 1
    Else
        Result = 0
        EqualFields = True
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Serialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMajor = .ReadProperty(PropMajor, DefaultMajor)
        mMinor = .ReadProperty(PropMinor, DefaultMinor)
        mBuild = .ReadProperty(PropBuild, DefaultBuild)
        mRevision = .ReadProperty(PropRevision, DefaultRevision)
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty PropMajor, mMajor
        .WriteProperty PropMinor, mMinor
        .WriteProperty PropBuild, mBuild
        .WriteProperty PropRevision, mRevision
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
'   IComparable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IComparable_CompareTo(Value As Variant) As Long
    IComparable_CompareTo = CompareTo(Value)
End Function

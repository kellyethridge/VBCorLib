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
' Gets the value of the major component of the version number for the current <a href="Version.html">Version</a> object.
'
' @return The major version number.
Public Property Get Major() As Long
    Major = mMajor
End Property

''
' Gets the value of the minor component of the version number for the current <a href="Version.html">Version</a> object.
'
' @return The minor version number.
Public Property Get Minor() As Long
    Minor = mMinor
End Property

''
' Gets the value of the build component of the version number for the current <a href="Version.html">Version</a> object.
'
' @return The build number, or -1 if the build number is undefined.
' @remarks For example, if the version number is 6.2.1.3, the build number is 1. If the version number is 6.2, the build number is undefined.
Public Property Get Build() As Long
    Build = mBuild
End Property

''
' Gets the value of the revision component of the version number for the current <a href="Version.html">Version</a> object.
'
' @return The revision number, or -1 if the revision number is undefined.
' @remarks For example, if the version number is 6.2.1.3, the revision number is 3. If the version number is 6.2, the revision number is undefined.
Public Property Get Revision() As Long
    Revision = mRevision
End Property

''
' Gets the low 16 bits of the revision number.
'
' @return A 16-bit signed integer.
Public Property Get MinorRevision() As Integer
    MinorRevision = AsWord(mRevision)
End Property

''
' Gets the high 16 bits of the revision number.
'
' @return A 16-bit signed integer.
Public Property Get MajorRevision() As Integer
    MajorRevision = Helper.ShiftRight(mRevision, 16)
End Property

''
' Returns a new <a href="Version.html">Version</a> object whose value is the same as the current <a href="Version.html">Version</a> object.
'
' @return A new object whose values are a copy of the current <a href="Version.html">Version</a> object.
Public Function Clone() As Version
    Set Clone = New Version
    Clone.InitClone mMajor, mMinor, mBuild, mRevision
End Function

''
' Compares the current <a href="Version.html">Version</a> object to a specified <a href="Version.html">Version</a> object and returns an indication of their relative values.
'
' @param value A <a href="Version.html">Version</a> object to compare to the current <a href="Version.html">Version</a> object, or <b>Nothing</b>.
' @return A signed integer that indicates the relative values of the two objects, as shown in the following table.
' <table style="width: auto">
'   <tr>
'       <th>Value</th>
'       <th>Meaning</th>
'   </tr>
'   <tr>
'        <td>Less than zero</td><td>The current <a href="Version.html">Version</a> object is a version before <i>Value</i>.</td>
'   </tr>
'    <tr>
'        <td>Zero</td><td>The current <a href="Version.html">Version</a> object is the same version as <i>Value</i>.</td>
'    </tr>
'    <tr>
'        <td>
'           Greater than zero</td><td>The current <a href="Version.html">Version</a> object is a version subsequent to <i>Value</i>.<br/>
'           -or-<br/>
'           <i>Value</i> is <b>Nothing</b>.
'        </td>
'    </tr>
' </table>
' @remarks The components of <b>Version</b> in decreasing order of importance are: major, minor, build, and revision. An unknown component is assumed to be older than any known component. For example:
' <ul>
'   <li>Version 1.1 is older than version 1.1.0.</li>
'   <li>Version 1.1 is older than version 1.1.1.</li>
'   <li>Version 1.1 is older than version 1.1.2.3.</li>
'   <li>Version 1.1.2 is older than version 1.1.2.4.</li>
'   <li>Version 1.2.5 is newer than version 1.2.3.4.</li>
' </ul>
Public Function CompareTo(ByVal Value As Version) As Long
    If Value Is Nothing Then
        CompareTo = 1
        Exit Function
    End If
    
    If Not EqualFields(mMajor, Value.Major, CompareTo) Then Exit Function
    If Not EqualFields(mMinor, Value.Minor, CompareTo) Then Exit Function
    If Not EqualFields(mBuild, Value.Build, CompareTo) Then Exit Function
    If Not EqualFields(mRevision, Value.Revision, CompareTo) Then Exit Function
End Function

''
' Converts the value of the current <a href="Version.html">Version</a> object to its equivalent string representation. A specified count indicates the number of components to return.
'
' @param FieldCount The number of components to return. The <i>FieldCount</i> ranges from 0 to 4.
' @return The string representation of the values of the major, minor, build, and revision components of the current <a href="Version.html">Version</a> object, each separated by a period character ('.'). The <i>FieldCount</i> parameter determines how many components are returned.
' <table style="width: auto">
'   <tr>
'       <th>FieldCount</th>
'       <th>Return Value</th>
'   </tr>
'   <tr>
'        <td>0</td><td>An empty string ("").</td>
'   </tr>
'    <tr>
'        <td>1</td><td>Major</td>
'    </tr>
'    <tr>
'        <td>2</td><td>Major.Minor</td>
'    </tr>
'    <tr>
'        <td>3</td><td>Major.Minor.Build</td>
'    </tr>
'    <tr>
'        <td>4</td><td>Major.Minor.Build.Revision</td>
'    </tr>
' </table>
' For example, if you create <a href="Version.html">Version</a> object using the constructor <code>Version(1,3,5)</code>, <code>ToString(2)</code> returns "1.3" and <code>ToString(4)</code> throws an exception.
' @exceptions
' <table style="width: auto">
'   <tr>
'       <th>Exception</th>
'       <th>Condition</th>
'   </tr>
'   <tr>
'        <td><a href="ArgumentException.html">ArgumentException</a></td>
'        <td>
'           <i>FieldCount</i> is less than zero, or more than four.<br/>
'           -or-<br/>
'           <i>FieldCount</i> is more than the number of components defined in the current <a href="Version.html">Version</a> object.
'        </td>
'   </tr>
' </table>
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
' Returns a value indicating whether the current <a href="Version.html">Version</a> object and a specified <a href="Version.html">Version</a> object represent the same value.
'
' @param value A <a href="Version.html">Version</a> object to compare to the current <a href="Version.html">Version</a> object, or <b>Nothing</b>.
' @return <b>True</b> if every component of the current <a href="Version.html">Version</a> object matches the corresponding component of the <i>Value</i> parameter; otherwise, <b>False</b>.
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
' Returns a hash code for the current <a href="Version.html">Version</a> object.
'
' @return A 32-bit signed integer hash code.
Public Function GetHashCode() As Long
    Dim Result As Long
    Result = Helper.ShiftLeft(mMajor, 24)
    Result = Result Or Helper.ShiftLeft(mMinor And &HFF, 16)
    Result = Result Or Helper.ShiftLeft(mBuild And &HFF, 8)
    GetHashCode = Result Or (mRevision And &HFF&)
End Function

''
' Determines whether two specified Version objects are equal.
'
' @param Value The <a href="Version.html">Version</a> object to test equality against.
' @return <b>True</b> if every component of the current <a href="Version.html">Version</a> object matches the corresponding component of the <i>Value</i> parameter; otherwise, <b>False</b>.
Public Function EqualTo(ByVal Value As Version) As Boolean
    EqualTo = (CompareTo(Value) = 0)
End Function

''
' Determines whether the current <a href="Version.html">Version</a> object is less than the specified <a href="Version.html">Version</a> object.
'
' @param Value The <a href="Version.html">Version</a> object to test the relationship against.
' @return <b>True</b> if the current <a href="Version.html">Version</a> object is less than <i>Value</i>; otherwise, <b>False</b>.
Public Function LessThan(ByVal Value As Version) As Boolean
    LessThan = (CompareTo(Value) < 0)
End Function

''
' Determines whether the current <a href="Version.html">Version</a> object is less than or equal to the specified <a href="Version.html">Version</a> object.
'
' @param Value The <a href="Version.html">Version</a> object to test the relationship against.
' @return <b>True</b> if the current <a href="Version.html">Version</a> object is less than or equal to <i>Value</i>; otherwise, <b>False</b>.
Public Function LessThanOrEqualTo(ByVal Value As Version) As Boolean
    LessThanOrEqualTo = (CompareTo(Value) <= 0)
End Function

''
' Determines whether the current <a href="Version.html">Version</a> object is greater than the specified <a href="Version.html">Version</a> object.
'
' @param Value The <a href="Version.html">Version</a> object to test the relationship against.
' @return <b>True</b> if the current <a href="Version.html">Version</a> object is greater than <i>Value</i>; otherwise, <b>False</b>.
Public Function GreaterThan(ByVal Value As Version) As Boolean
    GreaterThan = Not LessThanOrEqualTo(Value)
End Function

''
' Determines whether the current <a href="Version.html">Version</a> object is greater than or equal to the specified <a href="Version.html">Version</a> object.
'
' @param Value The <a href="Version.html">Version</a> object to test the relationship against.
' @return <b>True</b> if the current <a href="Version.html">Version</a> object is greater than or equal to <i>Value</i>; otherwise, <b>False</b>.
Public Function GreaterThanOrEqualTo(ByVal Value As Version) As Boolean
    GreaterThanOrEqualTo = Not LessThan(Value)
End Function

''
' Determines whether the current <a href="Version.html">Version</a> object is not equal to the specified <a href="Version.html">Version</a> object.
'
' @param Value The <a href="Version.html">Version</a> object to test the relationship against.
' @return <b>True</b> if the current <a href="Version.html">Version</a> object is not equal to <i>Value</i>; otherwise, <b>False</b>.
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

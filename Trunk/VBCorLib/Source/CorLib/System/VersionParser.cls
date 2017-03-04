VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Module: VersionParser
'
Option Explicit

Private Const MajorComponent As Long = 0
Private Const MinorComponent As Long = 1
Private Const BuildComponent As Long = 2
Private Const RevisionComponent As Long = 3

Private mException      As Exception
Private mComponents()   As String


Friend Sub ThrowError()
    Throw mException
End Sub

Friend Function TryParse(ByRef s As String, ByRef Result As Version) As Boolean
    Init s
    If Not HasValidComponentCount Then
        Exit Function
    End If
    
    Dim Major As Long
    Dim Minor As Long
    Dim Build As Long
    Dim Revision As Long
    
    If Not TryParseComponent(MajorComponent, Major) Then
        Exit Function
    End If
    If Not TryParseComponent(MinorComponent, Minor) Then
        Exit Function
    End If
    
    Select Case UBound(mComponents)
        Case 1
            Set Result = Cor.NewVersion(Major, Minor)
        Case 2
            If Not TryParseComponent(BuildComponent, Build) Then
                Exit Function
            End If
            Set Result = Cor.NewVersion(Major, Minor, Build)
        Case 3
            If Not (TryParseComponent(BuildComponent, Build) And TryParseComponent(RevisionComponent, Revision)) Then
                Exit Function
            End If
            Set Result = Cor.NewVersion(Major, Minor, Build, Revision)
    End Select
    
    TryParse = True
End Function

Private Sub Init(ByRef s As String)
    Set mException = Nothing
    mComponents = Split(s, ".")
End Sub

Private Function HasValidComponentCount() As Boolean
    Select Case UBound(mComponents)
        Case 1 To 3
            HasValidComponentCount = True
        Case Else
            Set mException = Cor.NewArgumentException(Environment.GetResourceString(Arg_VersionString))
    End Select
End Function

Private Function TryParseComponent(ByVal Component As Long, ByRef Result As Long) As Boolean
    On Error GoTo Catch
    
    Result = CLng(mComponents(Component))
    If Result < 0 Then
        Dim ParamName As String
        ParamName = Choose(Component + 1, "Major", "Minor", "Build", "Revision")
        mException = Cor.NewArgumentOutOfRangeException(ParamName, Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    Else
        TryParseComponent = True
    End If
    Exit Function
    
Catch:
    Select Case Err.Number
        Case vbTypeMismatch
            Set mException = Cor.NewFormatException(Environment.GetResourceString(Arg_VersionString))
        Case vbOverflow
            Set mException = New OverflowException
        Case Else
            Set mException = Cor.NewException(Err.Description, Err.Number)
    End Select
End Function










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

Private mException      As Exception
Private mComponents()   As String


Friend Property Get HasError() As Boolean
    HasError = Not mException Is Nothing
End Property

Friend Sub ThrowError()
    If HasError Then
        Throw mException
    End If
End Sub

Friend Function TryParse(ByRef s As String, ByRef Result As Version) As Boolean
    Init s
    If Not HasValidComponentCount Then
        Exit Function
    End If
    
    Select Case UBound(mComponents)
        Case 1
            Set Result = Cor.NewVersion(mComponents(0), mComponents(1))
        Case 2
            Set Result = Cor.NewVersion(mComponents(0), mComponents(1), CLng(mComponents(2)))
        Case 3
            Set Result = Cor.NewVersion(mComponents(0), mComponents(1), CLng(mComponents(2)), CLng(mComponents(3)))
    End Select
    
    TryParse = True
End Function

Private Sub Init(ByRef s As String)
    Set mException = Nothing
    mComponents = Split(s, ".")
End Sub

Private Function HasValidComponentCount() As Boolean
    Select Case UBound(mComponents)
        Case Is < 1, Is > 3
            Set mException = Cor.NewArgumentException(Environment.GetResourceString(Arg_VersionString))
        Case Else
            HasValidComponentCount = True
    End Select
End Function

Private Function TryParseComponent(ByVal Index As Long, ByRef Result As Long) As Boolean
    
End Function

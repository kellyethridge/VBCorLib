VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CaseInsensitiveComparerStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
' Module: CaseInsensitiveComparerStatic
'

''
' Provides access to default culture-sensitive <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> objects.
'
' @remarks This class cannot be instantiated. To access the defined properties use the class name.
' <code>Set c = CaseInsensitiveComparer.Default</code>
Option Explicit

Private mDefaultInvariant As CaseInsensitiveComparer

''
' Gets an instance of <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> that is associated with the <a href="Thread.Get.CurrentCulture.html">Thread.CurrentCulture</a> of the current thread and that is always available.
'
' @returns An instance of <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> that is associated with <a href="CultureInfo.Get.InvariantCulture.html">CultureInfo.InvariantCulture</a>.
Public Property Get Default() As CaseInsensitiveComparer
    Set Default = Cor.NewCaseInsensitiveComparer(CultureInfo.CurrentCulture)
End Property

''
' Gets an instance of <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> that is associated with <a href="CultureInfo.Get.InvariantCulture.html">CultureInfo.InvariantCulture</a> and that is always available.
'
' @return An instance of <a href="CaseInsensitiveComparer.html">CaseInsensitiveComparer</a> that is associated with <a href="CultureInfo.Get.InvariantCulture.html">CultureInfo.InvariantCulture</a>.
Public Property Get DefaultInvariant() As CaseInsensitiveComparer
    If mDefaultInvariant Is Nothing Then
        Set mDefaultInvariant = Cor.NewCaseInsensitiveComparer(CultureInfo.InvariantCulture)
    End If
    
    Set DefaultInvariant = mDefaultInvariant
End Property

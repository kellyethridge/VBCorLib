VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringComparerStatic"
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
' Module: StringComparerStatic
'

''
' Provides a set of <a href="StringComparer.html">StringComparer</a> objects implementing different string comparison rules.
'
' @see StringComparer
'
Option Explicit

Private mOrdinal            As NonculturalStringComparer
Private mOrdinalIgnoreCase  As NonculturalStringComparer
Private mBinaryCompare      As NonculturalStringComparer
Private mTextCompare        As NonculturalStringComparer
Private mInvariantCulture   As CultureAwareStringComparer
Private mInvariantCultureIgnoreCase As CultureAwareStringComparer


''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-sensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> object that is case-sensitive.
' @remarks The comparison is done using the Visual Basic <b>StrComp</b> function with <b>vbBinaryCompare</b> option.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.BinaryCompare
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get BinaryCompare() As StringComparer
    If mBinaryCompare Is Nothing Then
        Set mBinaryCompare = Cor.NewNonculturalStringComparer(StringComparison.BinaryCompare)
    End If
    
    Set BinaryCompare = mBinaryCompare
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-insensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> object that is case-insensitive.
' @remarks The comparison is done using the Visual Basic <b>StrComp</b> function with <b>vbTextCompare</b> option.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.TextCompare
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get TextCompare() As StringComparer
    If mTextCompare Is Nothing Then
        Set mTextCompare = Cor.NewNonculturalStringComparer(StringComparison.TextCompare)
    End If
    
    Set TextCompare = mTextCompare
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-sensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> object that is case-sensitive.
' @remarks The comparison is done in a similar fashion as using the standard equality check of two VisualBasic strings in a variant culture manner.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.Ordinal
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get Ordinal() As StringComparer
    If mOrdinal Is Nothing Then
        Set mOrdinal = Cor.NewNonculturalStringComparer(StringComparison.Ordinal)
    End If
    
    Set Ordinal = mOrdinal
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that provides case-insensitive string comparisons.
'
' @return A <a href="StringComparer.html">StringComparer</a> that is case-insensitive.
' @remarks The comparison is done like using the <b>StrComp</b> function to check two VisualBasic strings in a variant culture manner.
' <p>The <a href="StringComparer.html">StringComparer</a> object can be retrieved using the following code.</p>
' <pre>
'   Dim c As StringComparer
'   Set c = StringComparer.OrdinalIgnoreCase
' </pre>
' <p>The instance of the <a href="StringComparer.html">StringComparer</a> object returned is a private class and simply implements the <a href="StringComparer.html">StringComparer</a> interface.</p>
Public Property Get OrdinalIgnoreCase() As StringComparer
    If mOrdinalIgnoreCase Is Nothing Then
        Set mOrdinalIgnoreCase = Cor.NewNonculturalStringComparer(StringComparison.OrdinalIgnoreCase)
    End If
    
    Set OrdinalIgnoreCase = mOrdinalIgnoreCase
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that performs a case-sensitive string comparison using the word comparison rules of the invariant culture.
'
' @return A new <a href="StringComparer.html">StringComparer</a> object.
' @remarks The <a href="StringComparer.html">StringComparer</a> returned by the <b>InvariantCulture</b> property
' compares strings in a linguistically relevant manner, but it is not suitable for display in any particular culture.
' Its major application is to order strings in a way that will be identical across cultures.
' <p>The invariant culture is the <a href="CultureInfo.html">CultureInfo</a> object returned by the <a href="CultureInfo.Get.InvariantCulture.html">CultureInfo.InvariantCulture</a> property.</p>
' <p>The <a href="StringComparer.html">StringComparer</a> can be accessed using the following code.</p>
' <code>Set c = StringComparer.InvariantCulture</code>
Public Property Get InvariantCulture() As StringComparer
    If mInvariantCulture Is Nothing Then
        Set mInvariantCulture = Cor.NewCultureAwareStringComparer(CultureInfo.InvariantCulture, False)
    End If
    
    Set InvariantCulture = mInvariantCulture
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that performs a case-insensitive string comparison using the word comparison rules of the invariant culture.
'
' @return A new <a href="StringComparer.html">StringComparer</a> object.
' @remarks The <a href="StringComparer.html">StringComparer</a> returned by the <b>InvariantCultureIgnoreCase</b> property
' compares strings in a linguistically relevant manner that ignores case, but it is not suitable for display in any particular culture.
' Its major application is to order strings in a way that will be identical across cultures.
' <p>The invariant culture is the <a href="CultureInfo.html">CultureInfo</a> object returned by the <a href="CultureInfo.Get.InvariantCulture.html">CultureInfo.InvariantCulture</a> property.</p>
' <p>The <a href="StringComparer.html">StringComparer</a> can be accessed using the following code.</p>
' <code>Set c = StringComparer.InvariantCultureIgnoreCase</code>
Public Property Get InvariantCultureIgnoreCase() As StringComparer
    If mInvariantCultureIgnoreCase Is Nothing Then
        Set mInvariantCultureIgnoreCase = Cor.NewCultureAwareStringComparer(CultureInfo.InvariantCulture, True)
    End If
    
    Set InvariantCultureIgnoreCase = mInvariantCultureIgnoreCase
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that performs a case-sensitive string comparison using the word comparison rules of the current culture.
'
' @return A new <a href="StringComparer.html">StringComparer</a> object.
' @remarks The <a href="StringComparer.html">StringComparer</a> returned by the <b>CurrentCulture</b> property can be
' used when strings are linguistically relevant. For example, if strings are displayed to the user, or if strings are
' the result of user interaction, culture-sensitive string comparison should be used to order the string data.
' <p>The current culture is the <a href="CultureInfo.html">CultureInfo</a> object associated with the current thread.</p>
' <p>The <a href="StringComparer.html">StringComparer</a> can be accessed using the following code.</p>
' <code>Set c = StringComparer.CurrentCulture</code>
' <p>Each call to the <b>CurrentCulture</b> property get accessor returns a new <a href="StringComparer.html">StringComparer</a> object.</p>
Public Property Get CurrentCulture() As StringComparer
    Set CurrentCulture = Cor.NewCultureAwareStringComparer(CultureInfo.CurrentCulture, False)
End Property

''
' Gets a <a href="StringComparer.html">StringComparer</a> object that performs a case-insensitive string comparison using the word comparison rules of the current culture.
'
' @return A new <a href="StringComparer.html">StringComparer</a> object.
' @remarks The current culture is the <a href="CultureInfo.html">CultureInfo</a> object associated with the current thread.
' <p>The <a href="StringComparer.html">StringComparer</a> returned by the <b>CurrentCultureIgnoreCase</b> property can be
' used when strings are linguistically relevant but their case is not. For example, if strings are displayed to the user
' but case is unimportant, culture-sensitive, case-insensitive string comparison should be used to order the string data.</p>
' <code>Set c = StringComparer.CurrentCultureIgnoreCase</code>
' <p>Each call to the <b>CurrentCultureIgnoreCase</b> property get accessor returns a new <a href="StringComparer.html">StringComparer</a> object.</p>
Public Property Get CurrentCultureIgnoreCase() As StringComparer
    Set CurrentCultureIgnoreCase = Cor.NewCultureAwareStringComparer(CultureInfo.CurrentCulture, True)
End Property


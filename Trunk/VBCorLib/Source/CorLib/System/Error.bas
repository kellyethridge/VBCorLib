Attribute VB_Name = "Error"
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
' Module: Error
'
Option Explicit

Public Function Argument(ByVal Message As ArgumentString, Optional ByVal Parameter As ParameterString = Parameter_None) As ArgumentException
    Set Argument = Cor.NewArgumentException(Environment.GetResourceString(Message), Environment.GetResourceString(Parameter))
End Function

Public Function ArgumentNull(ByVal Parameter As ParameterString, Optional ByVal Message As ArgumentNullString = ArgumentNull_Exception) As ArgumentNullException
    Set ArgumentNull = Cor.NewArgumentNullException(Environment.GetResourceString(Parameter), Environment.GetResourceString(Message))
End Function

Public Function ArgumentOutOfRange(ByVal Parameter As ParameterString, Optional ByVal Message As ArgumentOutOfRangeString = ArgumentOutOfRange_Exception) As ArgumentOutOfRangeException
    Set ArgumentOutOfRange = Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(Parameter), Environment.GetResourceString(Message))
End Function

Public Function Rank(Optional ByVal Message As ResourceString = Rank_MultiDimNotSupported) As RankException
    Set Rank = Cor.NewRankException(Environment.GetResourceString(Message))
End Function

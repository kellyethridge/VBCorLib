Attribute VB_Name = "TimeConstants"
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
' Module: TimeConstants
'
Option Explicit

Public Const TicksPerMillisecond    As Long = 10000
Public Const SecondsPerMinute       As Long = 60
Public Const MinutesPerHour         As Long = 60
Public Const HoursPerDay            As Long = 24
Public Const DaysPerYear            As Long = 365
Public Const MonthsPerYear          As Long = 12
Public Const DaysPer4Years          As Long = DaysPerYear * 4 + 1

Public Const MillisecondsPerTick    As Currency = 0.0001@
Public Const MillisecondsPerSecond  As Long = 1000
Public Const MillisecondsPerMinute  As Currency = MillisecondsPerSecond * SecondsPerMinute
Public Const MillisecondsPerHour    As Currency = MillisecondsPerMinute * MinutesPerHour
Public Const MillisecondsPerDay     As Currency = MillisecondsPerHour * HoursPerDay

Public Const DaysTo1899             As Long = 693593
Public Const MillisecondsTo1899     As Currency = 1@ * DaysTo1899 * MillisecondsPerDay

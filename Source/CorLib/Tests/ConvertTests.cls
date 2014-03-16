VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConvertTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ConvertTests
'
Option Explicit
Implements ITestCaseSource

Private Const AllByteValuesAsBase64 As String = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="
Private Const AllByteValuesAsBase64WithLineBreaks As String = "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4" & vbCrLf & _
                                                              "OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3Bx" & vbCrLf & _
                                                              "cnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmq" & vbCrLf & _
                                                              "q6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj" & vbCrLf & _
                                                              "5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/w=="

Private AllByteValues() As Byte

Public Sub FixtureSetup()
    Set Thread.CurrentThread.CurrentCulture = Cor.NewCultureInfo("en-US", False)
    
    ReDim AllByteValues(0 To 255)
    Dim i As Long
    For i = 0 To 255
        AllByteValues(i) = i
    Next
End Sub

Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBComp.TestCaseCollector)
    Select Case Test.MethodName
        Case "ToBase64String_WithValues_ReturnsExpected"
            Test.Use "", ""
            Test.Use "f", "Zg=="
            Test.Use "fo", "Zm8="
            Test.Use "foo", "Zm9v"
            Test.Use "foob", "Zm9vYg=="
            Test.Use "fooba", "Zm9vYmE="
            Test.Use "foobar", "Zm9vYmFy"
            
        Case "FromBase64String_WithValues_ReturnsExpected"
            Test.Use "", ""
            Test.Use "Zg==", "f"
            Test.Use "Zm8=", "fo"
            Test.Use "Zm9v", "foo"
            Test.Use "Zm9vYg==", "foob"
            Test.Use "Zm9vYmE=", "fooba"
            Test.Use "Zm9vYmFy", "foobar"
            
        Case "ToString_WithValuesUsingDefaultFormatting_ReturnsExpected"
            ' String
            Test.Use "", ""
            Test.Use "a", "a"
            
            ' Byte
            Test.Use CByte(0), "0"
            Test.Use CByte(1), "1"
            Test.Use CByte(255), "255"
            
            ' Integer
            Test.Use 0, "0"
            Test.Use 1, "1"
            Test.Use CInt(32767), "32767"
            Test.Use -1, "-1"
            Test.Use CInt(-32768), "-32768"
            
            ' Long
            Test.Use 0&, "0"
            Test.Use 1&, "1"
            Test.Use -1&, "-1"
            Test.Use &H7FFFFFFF, "2147483647"
            Test.Use &H80000000, "-2147483648"
            
            ' Double
            Test.Use 0#, "0"
            Test.Use 1#, "1"
            Test.Use 0.1, "0.1"
            Test.Use 0.001, "0.001"
            Test.Use 100.001, "100.001"
            Test.Use 100, "100"
            Test.Use 123456789012345#, "123456789012345"
            Test.Use -123456789012345#, "-123456789012345"
            Test.Use 1.79769313486231E+308, "1.79769313486231E+308"
            Test.Use 4.94065645841247E-324, "4.94065645841247E-324"
            Test.Use -1.79769313486231E+308, "-1.79769313486231E+308"
            Test.Use -4.94065645841247E-324, "-4.94065645841247E-324"
            Test.Use MakeImaginary(0), "NaN"
            Test.Use MakeImaginary(1), "Infinity"
            Test.Use MakeImaginary(-1), "-Infinity"
            
            ' Single
            Test.Use 0!, "0"
            Test.Use 0.1!, "0.1"
            Test.Use 1!, "1"
            Test.Use 0.001, "0.001"
            Test.Use 100.001, "100.001"
            Test.Use 100, "100"
            Test.Use 1234567!, "1234567"
            Test.Use -1234567!, "-1234567"
            Test.Use 1.401298E-45, "1.401298E-45"
            Test.Use 3.402823E+38, "3.402823E+38"
            Test.Use -1.401298E-45, "-1.401298E-45"
            Test.Use -3.402823E+38, "-3.402823E+38"
            
            ' Currency
            Test.Use 0@, "0"
            Test.Use 0.1@, "0.1"
            Test.Use -0.1@, "-0.1"
            Test.Use 922337203685477.5807@, "922337203685477.5807"
            Test.Use -922337203685477.5807@ - 0.0001@, "-922337203685477.5808"
            
            ' Decimal
            Test.Use CDec(0), "0"
            Test.Use CDec(0.1), "0.1"
            Test.Use CDec("79228162514264337593543950335"), "79228162514264337593543950335"
            Test.Use CDec("-79228162514264337593543950335"), "-79228162514264337593543950335"
            Test.Use CDec("7.9228162514264337593543950335"), "7.9228162514264337593543950335"
            Test.Use CDec("-7.9228162514264337593543950335"), "-7.9228162514264337593543950335"
            Test.Use CDec("0.0000000000000000000000000001"), "0.0000000000000000000000000001"
            Test.Use CDec("-0.0000000000000000000000000001"), "-0.0000000000000000000000000001"
            
            ' Date
            Test.Use #12:00:00 AM#, "12/30/1899 12:00:00 AM"
            Test.Use #1/1/100#, "1/1/0100 12:00:00 AM"
            Test.Use #12/31/9999 11:59:59 PM#, "12/31/9999 11:59:59 PM"
            
            ' Boolean
            Test.Use True, "True"
            Test.Use False, "False"
            
            ' Object
            Dim Formattable As New MockFormattable
            Formattable.Result = "abc"
            Test.Use Formattable, "abc"
            Dim Object As New MockObject
            Object.ToStringResult = "123"
            Test.Use Object, "123"
            Test.Use Nothing, ""
            
            ' VBA
            Test.Use Empty, "Empty"
            Test.Use Null, "Null"
            
        Case "ToString_WithLongsUsingFormatting_ReturnsExpected"
            ' General Format
            Test.Use 0, "", "0"
            Test.Use 1, "", "1"
            Test.Use -1, "", "-1"
            Test.Use 100, "", "100"
            Test.Use &H7FFFFFFF, "", "2147483647"
            Test.Use &H80000000, "", "-2147483648"
            Test.Use 0, "g", "0"
            Test.Use 1, "g", "1"
            Test.Use -1, "g", "-1"
            Test.Use &H7FFFFFFF, "g", "2147483647"
            Test.Use &H80000000, "g", "-2147483648"
            Test.Use 0, "g0", "0"
            Test.Use 1, "g0", "1"
            Test.Use -1, "g0", "-1"
            Test.Use &H7FFFFFFF, "g0", "2147483647"
            Test.Use &H80000000, "g0", "-2147483648"
            Test.Use 0, "g1", "0"
            Test.Use 1, "g1", "1"
            Test.Use -1, "g1", "-1"
            Test.Use &H7FFFFFFF, "g1", "2e+09"
            Test.Use &H80000000, "g1", "-2e+09"
            Test.Use 0, "g5", "0"
            Test.Use 1, "g5", "1"
            Test.Use -1, "g5", "-1"
            Test.Use &H7FFFFFFF, "g5", "2.1475e+09"
            Test.Use &H80000000, "g5", "-2.1475e+09"
            Test.Use 12345, "ggg", "ggg"
            
            ' Currency Format
            Test.Use 1, "c", "$1.00"
            Test.Use 1, "c0", "$1"
            Test.Use 1, "c1", "$1.0"
            Test.Use 1, "c3", "$1.000"
            Test.Use -1, "c", "($1.00)"
            Test.Use 12345, "c", "$12,345.00"
            
            ' Decimal Format
            Test.Use 1, "d", "1"
            Test.Use 1, "d1", "1"
            Test.Use 1, "d2", "01"
            Test.Use -1, "d2", "-01"

            ' Exponent Format
            Test.Use 1, "e", "1.000000e+000"
            Test.Use 1, "e0", "1e+000"
            Test.Use 1, "e1", "1.0e+000"
            Test.Use 10, "e", "1.000000e+001"
            Test.Use 10, "e0", "1e+001"
            Test.Use 10, "e1", "1.0e+001"
            Test.Use -1, "e", "-1.000000e+000"
            Test.Use 19, "e0", "2e+001"
            Test.Use 12345, "e", "1.234500e+004"
            Test.Use -12345, "e4", "-1.2345e+004"
            Test.Use 12345, "e3", "1.235e+004"
            Test.Use 12345, "E1", "1.2E+004"
            Test.Use 0, "e", "0.000000e+000"
            Test.Use 123456789, "e4", "1.2346e+008"
            
            ' Fixed Format
            Test.Use 0, "f", "0.00"
            Test.Use 1, "f", "1.00"
            Test.Use 100, "f", "100.00"
            Test.Use 1234, "f", "1234.00"
            Test.Use -1234, "f", "-1234.00"
            Test.Use 1234, "f4", "1234.0000"
            Test.Use 1234, "f0", "1234"
            
            ' Number Format
            Test.Use 0, "n", "0.00"
            Test.Use 1, "n", "1.00"
            Test.Use 100, "n", "100.00"
            Test.Use 1234, "n", "1,234.00"
            Test.Use 1234, "n0", "1,234"
            Test.Use -1234, "n", "-1,234.00"
            Test.Use 1234, "n4", "1,234.0000"
            Test.Use 123, "n", "123.00"
            
            ' Percentage Format
            Test.Use 0, "p", "0.00 %"
            Test.Use 1, "p", "100.00 %"
            Test.Use 156, "p", "15,600.00 %"
            
            ' Hex Format
            Test.Use 0, "X", "0"
            Test.Use 0, "x2", "00"
            Test.Use 1, "x", "1"
            Test.Use 1, "x2", "01"
            Test.Use 255, "x", "ff"
            Test.Use 255, "X", "FF"
            Test.Use 255, "X4", "00FF"
            Test.Use -1, "x", "ffffffff"

        Case "ToString_WithDoubleUsingFormatting_ReturnsExpected"
            ' General Format
            Test.Use 0, "g", "0"
            Test.Use 0, "g1", "0"
            Test.Use 0, "g5", "0"
            Test.Use 1, "g", "1"
            Test.Use 1, "g0", "1"
            Test.Use 1, "g1", "1"
            Test.Use 1, "g2", "1"
            Test.Use 0.1, "g", "0.1"
            Test.Use 0.1, "g0", "0.1"
            Test.Use 0.1, "g1", "0.1"
            Test.Use 0.1, "g2", "0.1"
            Test.Use 10, "g", "10"
            Test.Use 10, "g0", "10"
            Test.Use 10, "g1", "1e+01"
            Test.Use 10, "g2", "10"
            Test.Use 123, "g", "123"
            Test.Use 123, "g0", "123"
            Test.Use 123, "g1", "1e+02"
            Test.Use 123, "g2", "1.2e+02"
            Test.Use 123, "g3", "123"
            Test.Use 123000, "g", "123000"
            Test.Use 123000, "g0", "123000"
            Test.Use 123000, "g1", "1e+05"
            Test.Use 123000, "g2", "1.2e+05"
            Test.Use 123000, "g3", "1.23e+05"
            Test.Use 123000, "g4", "1.23e+05"
            Test.Use 0.000123, "g", "0.000123"
            Test.Use 0.000123, "g0", "0.000123"
            Test.Use 0.000123, "g1", "0.0001"
            Test.Use 0.000123, "g2", "0.00012"
            Test.Use 0.000123, "g3", "0.000123"
            Test.Use 0.000123, "g4", "0.000123"
            Test.Use 100.001, "g", "100.001"
            Test.Use 100.001, "g0", "100.001"
            Test.Use 100.001, "g1", "1e+02"
            Test.Use 100.001, "g2", "1e+02"
            Test.Use 100.001, "g3", "100"
            Test.Use 100.001, "g4", "100"
            Test.Use 100.001, "g5", "100"
            Test.Use 100.001, "g6", "100.001"
            Test.Use 1.2345E+27, "g28", "1234500000000000000000000000"
            Test.Use 12345, "g4", "1.235e+04"
            
            ' Currency Format
            Test.Use 0, "c", "$0.00"
            Test.Use 0, "c0", "$0"
            Test.Use 0, "c1", "$0.0"
            Test.Use 0, "c2", "$0.00"
            Test.Use 0, "c3", "$0.000"
            Test.Use 1000, "c", "$1,000.00"
            Test.Use 1000, "c0", "$1,000"
            Test.Use 1000, "c1", "$1,000.0"
            Test.Use 1000, "c2", "$1,000.00"
            Test.Use 1000, "c3", "$1,000.000"
            Test.Use 1.23, "c", "$1.23"
            Test.Use 1234.56, "c", "$1,234.56"
            Test.Use 1234567.89, "c", "$1,234,567.89"
            Test.Use 12345678.91, "c", "$12,345,678.91"
            Test.Use 123.45, "c0", "$123"
            Test.Use 123.45678, "c8", "$123.45678000"
            Test.Use -123.45, "c0", "($123)"
            Test.Use 0.00123, "c", "$0.00"
            Test.Use 0.001235, "c5", "$0.00124"
            Test.Use 123#, "c", "$123.00"
            Test.Use 1.23123123E+20, "c5", "$123,123,123,000,000,000,000.00000"
            
            ' Scientific Format
            Test.Use 12345#, "e", "1.234500e+004"
            Test.Use -12345#, "e4", "-1.2345e+004"
            Test.Use 12345#, "e3", "1.235e+004"
            Test.Use 12345#, "e0", "1e+004"
            Test.Use 12345#, "E1", "1.2E+004"
            Test.Use 1234500#, "e", "1.234500e+006"
            Test.Use 1234500#, "e3", "1.235e+006"
            Test.Use 0.0012345, "e3", "1.235e-003"
            Test.Use 0.0012345, "e", "1.234500e-003"
            Test.Use 1.2345E-200, "e7", "1.2345000e-200"
            Test.Use 1#, "e", "1.000000e+000"
            Test.Use 1#, "e1", "1.0e+000"
            Test.Use 0#, "e", "0.000000e+000"
            Test.Use 0.1234, "e3", "1.234e-001"
            
            ' Fixed Format
            Test.Use 12345#, "f", "12345.00"
            Test.Use 12345#, "f0", "12345"
            Test.Use 123450#, "f10", "123450.0000000000"
            Test.Use 123.45, "f", "123.45"
            Test.Use 1.2345, "f", "1.23"
            Test.Use 1.2345, "f3", "1.235"
            Test.Use 0.12345, "f", "0.12"
            Test.Use 0.00012345, "f3", "0.000"
            Test.Use 0.00012345, "f7", "0.0001235"
            Test.Use 0.00012345, "f20", "0.00012345000000000000"
            Test.Use 1#, "f", "1.00"
            Test.Use 0#, "f", "0.00"
            
            ' Number Format
            Test.Use 1.23, "n", "1.23"
            Test.Use 1234.56, "n", "1,234.56"
            Test.Use 1234567.89, "n", "1,234,567.89"
            Test.Use 12345678.91, "n", "12,345,678.91"
            Test.Use 123.45, "n0", "123"
            Test.Use 123.45678, "n8", "123.45678000"
            Test.Use -123.45, "n0", "-123"
            Test.Use 0.00123, "n", "0.00"
            Test.Use 0.001235, "n5", "0.00124"
            Test.Use 123#, "n", "123.00"
            Test.Use 1.23123123E+20, "n5", "123,123,123,000,000,000,000.00000"
            Test.Use 0#, "n", "0.00"
            
            ' Percent Format
            Test.Use 0#, "p", "0.00 %"
            Test.Use 1.23, "p", "123.00 %"
            Test.Use 1234.56, "p", "123,456.00 %"
            Test.Use 1234567.89, "p", "123,456,789.00 %"
            Test.Use 12345678.91, "p", "1,234,567,891.00 %"
            Test.Use 123.45, "p0", "12,345 %"
            Test.Use 123.45678, "p8", "12,345.67800000 %"
            Test.Use -123.45, "p0", "-12,345 %"
            Test.Use 0.00123, "p", "0.12 %"
            Test.Use 0.001235, "p3", "0.124 %"
            Test.Use 123#, "p", "12,300.00 %"
            Test.Use 1.23123123E+20, "p5", "12,312,312,300,000,000,000,000.00000 %"

        Case "ToString_WithSingleUsingFormatting_ReturnsExpected"
            ' General Format
            Test.Use 12345!, "", "12345"
            Test.Use 12345!, "g", "12345"
            Test.Use -12345!, "g", "-12345"
            Test.Use 1.2345!, "g", "1.2345"
            Test.Use 0.012345!, "g", "0.012345"
            Test.Use 1234.568!, "g", "1234.568"
            Test.Use 1.2345E+25!, "g26", "12345000400000000000000000"
            Test.Use 1.2345E-05!, "", "1.2345E-05"
            Test.Use 1.2345E+15!, "", "1.2345E+15"
            Test.Use 1.2345E-05!, "G", "1.2345E-05"
            Test.Use 1.2345E+15!, "G", "1.2345E+15"
            Test.Use 12345!, "g4", "1.235e+04"
            
            ' Fixed Format
            Test.Use 12345!, "f", "12345.00"
            Test.Use 12345!, "f0", "12345"
            Test.Use 123450!, "f10", "123450.0000000000"
            Test.Use 123.45!, "f", "123.45"
            Test.Use 1.2345!, "f", "1.23"
            Test.Use 1.2345!, "f3", "1.235"
            Test.Use 0.12345!, "f", "0.12"
            Test.Use 1.2345E-04!, "f3", "0.000"
            Test.Use 1.2345E-04!, "f7", "0.0001235"
            Test.Use 1.2345E-04!, "f20", "0.00012345000000000000"

        Case "ToString_WithNumberNegativePattern_ReturnsExpected"
            Test.Use -1, NumberNegativePatterns.[- n], "- 1.00"
            Test.Use -1, NumberNegativePatterns.[(n)], "(1.00)"
            Test.Use -1, NumberNegativePatterns.[-n], "-1.00"
            Test.Use -1, NumberNegativePatterns.[n-], "1.00-"
            Test.Use -1, NumberNegativePatterns.[n -], "1.00 -"
            
        Case "ToString_WithCurrencyNegativePattern_ReturnsExpected"
            Test.Use -1, CurrencyNegativePatterns.[-$ n], "-$ 1.00"
            Test.Use -1, CurrencyNegativePatterns.[$ -n], "$ -1.00"
            Test.Use -1, CurrencyNegativePatterns.[$ n-], "$ 1.00-"
            Test.Use -1, CurrencyNegativePatterns.[-$n], "-$1.00"
            Test.Use -1, CurrencyNegativePatterns.[$-n], "$-1.00"
            Test.Use -1, CurrencyNegativePatterns.[$n-], "$1.00-"
            Test.Use -1, CurrencyNegativePatterns.[($ n)], "($ 1.00)"
            Test.Use -1, CurrencyNegativePatterns.[($n)], "($1.00)"
            Test.Use -1, CurrencyNegativePatterns.[(n $)], "(1.00 $)"
            Test.Use -1, CurrencyNegativePatterns.[(n$)], "(1.00$)"
            Test.Use -1, CurrencyNegativePatterns.[-n $], "-1.00 $"
            Test.Use -1, CurrencyNegativePatterns.[n- $], "1.00- $"
            Test.Use -1, CurrencyNegativePatterns.[n $-], "1.00 $-"
            Test.Use -1, CurrencyNegativePatterns.[-n$], "-1.00$"
            Test.Use -1, CurrencyNegativePatterns.[n-$], "1.00-$"
            Test.Use -1, CurrencyNegativePatterns.[n$-], "1.00$-"
            
        Case "ToString_WithCurrencyPositivePattern_ReturnsExpected"
            Test.Use 1, CurrencyPositivePatterns.[$ n], "$ 1.00"
            Test.Use 1, CurrencyPositivePatterns.[$n], "$1.00"
            Test.Use 1, CurrencyPositivePatterns.[n $], "1.00 $"
            Test.Use 1, CurrencyPositivePatterns.[n$], "1.00$"
            
        Case "ToString_WithPercentNegativePattern_ReturnsExpected"
            Test.Use -1, PercentNegativePatterns.[-# %], "-100.00 %"
            Test.Use -1, PercentNegativePatterns.[-#%], "-100.00%"
            Test.Use -1, PercentNegativePatterns.[-% n], "-% 100.00"
            Test.Use -1, PercentNegativePatterns.[% -n], "% -100.00"
            Test.Use -1, PercentNegativePatterns.[% n-], "% 100.00-"
            Test.Use -1, PercentNegativePatterns.[-%#], "-%100.00"
            Test.Use -1, PercentNegativePatterns.[%-n], "%-100.00"
            Test.Use -1, PercentNegativePatterns.[%n-], "%100.00-"
            Test.Use -1, PercentNegativePatterns.[n- %], "100.00- %"
            Test.Use -1, PercentNegativePatterns.[n %-], "100.00 %-"
            Test.Use -1, PercentNegativePatterns.[n-%], "100.00-%"
            Test.Use -1, PercentNegativePatterns.[n%-], "100.00%-"
            
        Case "ToString_WithCustomCurrencyGroupSizes_ReturnsCustomGrouping"
            Test.Use 12345678.91, NewLongs(1, 2, 3), "$12,345,67,8.91"
            Test.Use 123456789.12, NewLongs(1, 2, 1, 0), "$12345,6,78,9.12"
            Test.Use 1234567890#, NewLongs(), "$1234567890.00"
            
        Case "ToString_WithCustomNumberGroupSizes_ReturnsCustomGrouping"
            Test.Use 12345678.91, NewLongs(1, 2, 3), "12,345,67,8.91"
            Test.Use 123456789.12, NewLongs(1, 2, 1, 0), "12345,6,78,9.12"
            Test.Use 1234567890#, NewLongs(), "1234567890.00"
            
        Case "ToString_WithCustomPercentageGroupSizes_ReturnsCustomGrouping"
'    Info.PercentGroupSizes = NewLongs(1, 2, 3)
'    Assert.AreEqual "1,234,567,89,1.00 %", Info.Format(12345678.91, "p")
'
'    Info.PercentGroupSizes = NewLongs(1, 2, 1, 0)
'    Assert.AreEqual "1234567,8,91,2.00 %", Info.Format(123456789.12, "p")
'
'    Info.PercentGroupSizes = cArray.CreateInstance(vbLong)
'    Assert.AreEqual "123456789000 %", Info.Format(1234567890#, "p0")
            Test.Use 12345678.91, NewLongs(1, 2, 3), "1,234,567,89,1.00 %"
            Test.Use 123456789.12, NewLongs(1, 2, 1, 0), "1234567,8,91,2.00 %"
            Test.Use 1234567890#, NewLongs(), "123456789000.00 %"
    End Select
End Sub

Public Sub ToBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual As String
    Dim Bytes() As Byte
    Bytes = StrConv(Source, vbFromUnicode)
    
    Actual = Convert.ToBase64String(Bytes)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub FromBase64String_WithValues_ReturnsExpected(ByVal Source As String, ByVal Expected As String)
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(Source)
    
    Dim Bytes() As Byte
    Bytes = StrConv(Expected, vbFromUnicode)
    Assert.That Actual, Iz.EqualTo(Bytes)
End Sub

Public Sub ToBase64String_WithAllByteValues_ReturnsAllByteValuesAsBase64()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
End Sub

Public Sub ToBase64String_WithAllByteValuesAndInsertLineBreaks_ReturnsAllByteValuesAsBase64WithLineBreaks()
    Dim Actual As String
    
    Actual = Convert.ToBase64String(AllByteValues, InsertLineBreaks:=True)
    
    Assert.That Actual, Iz.EqualTo(AllByteValuesAsBase64WithLineBreaks)
End Sub

Public Sub FromBase64String_WithAllByteValuesAsBase64WithLineBreaks_ReturnsAllByteValues()
    Dim Actual() As Byte
    
    Actual = Convert.FromBase64String(AllByteValuesAsBase64WithLineBreaks)
    
    Assert.That Actual, Iz.EqualTo(AllByteValues)
End Sub

Public Sub ToBase64String_WithNullArray_Throws()
    Dim Bytes() As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64String_WithIndexLessThanLBound_Throws()
    Dim Bytes(1 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 1
    
    AssertArgumentOutOfRangeException Err, "Index"
End Sub

Public Sub ToBase64String_WithLengthLessThanZero_Throws()
    Dim Bytes(0) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64String_WithIndexAndLengthGreaterThanUBound_Throws()
    Dim Bytes(0 To 2) As Byte
    
    On Error Resume Next
    Convert.ToBase64String Bytes, 0, 4
    
    AssertArgumentOutOfRangeException Err
End Sub

Public Sub ToBase64CharArray_WithNullInArray_Throws()
    Dim InArray() As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub ToBase64CharArray_WithNullOutArray_Throws()
    Dim InArray(0) As Byte
    Dim OutArray() As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentNullException Err, "OutArray"
End Sub

Public Sub ToBase64CharArray_WithIndexInLessThanLBound_Throws()
    Dim InArray(1 To 2) As Byte
    Dim OutArray(0) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 0, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub ToBase64CharArray_WithIndexOutLessThanLBound_Throws()
    Dim InArray(2) As Byte
    Dim OutArray(1 To 5) As Integer
    
    On Error Resume Next
    
    Convert.ToBase64CharArray InArray, 0, 1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexOut"
End Sub

Public Sub ToBase64CharArray_WithLengthLessThanZero_Throws()
    Dim InArray(1) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, -1, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub ToBase64CharArray_WithIndexInAndLengthGreaterThanUBound_Throws()
    Dim InArray(1) As Byte
    Dim OutArray(2) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub ToBase64CharArray_WithIndexOutAndLengthGreaterThanUBound_Throws()
    Dim InArray(2) As Byte
    Dim OutArray(1) As Integer
    
    On Error Resume Next
    Convert.ToBase64CharArray InArray, 0, 3, OutArray, 0
    
    AssertArgumentOutOfRangeException Err, "IndexOut"
End Sub

Public Sub FromBase64CharArray_WithNullArray_Throws()
    Dim InArray() As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 0
    
    AssertArgumentNullException Err, "InArray"
End Sub

Public Sub FromBase64CharArray_WithIndexInLessThanLBound_Throws()
    Dim InArray(1 To 2) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 0
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub FromBase64CharArray_WithLengthLessThanZero_Throws()
    Dim InArray(0) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, -1
    
    AssertArgumentOutOfRangeException Err, "Length"
End Sub

Public Sub FromBase64CharArray_WithIndexInAndLengthGreaterThanUBound_Throws()
    Dim InArray(2) As Integer
    
    On Error Resume Next
    Convert.FromBase64CharArray InArray, 0, 4
    
    AssertArgumentOutOfRangeException Err, "IndexIn"
End Sub

Public Sub FromBase64String_WithMalformedString_Throws()
    Const Malformed As String = "Zm9vYg=" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub FromBase64String_WithTooMuchPadding_Throws()
    Const Malformed As String = "Zm9vYg===" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub FromBase64String_WithInvalidCharater_Throws()
    Const Malformed As String = "Zm9vY.==" ' correct form: Zm9vYg==
    
    On Error Resume Next
    Convert.FromBase64String Malformed
    
    AssertFormatException Err
End Sub

Public Sub ToString_WithValuesUsingDefaultFormatting_ReturnsExpected(ByRef Value As Variant, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithFormattable_SuppliesFormatInformation()
    Dim Mock As New MockFormattable
    
    Convert.ToString Mock, "ab"
    
    Dim Actual As String
    Actual = Mock.Format
    Assert.That Actual, Iz.EqualTo("ab")
End Sub

Public Sub ToString_WithFormattable_SuppliesProvider()
    Dim Mock As New MockFormattable
    Dim Provider As New NumberFormatInfo
    
    Convert.ToString Mock, , Provider
    
    Dim Actual As IFormatProvider
    Set Actual = Mock.Provider
    Assert.That Actual, Iz.SameAs(Provider)
End Sub

Public Sub ToString_WithMissingValue_ReturnsEmptyString()
    Dim Actual As String
    
    Actual = Convert.ToString(Missing)
    
    Assert.That Actual, Iz.Empty
End Sub

Public Sub ToString_WithLongsUsingFormatting_ReturnsExpected(ByVal Value As Long, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithLongUsingModifiedGroupSize_ReturnsExpected()
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.CurrencyGroupSizes = NewLongs(1, 2, 3)
    
    Actual = Convert.ToString(12345678, "c", Info)
    
    Assert.That Actual, Iz.EqualTo("$12,345,67,8.00")
End Sub

Public Sub ToString_WithLongUsingModifiedGroupSize_ReturnsExpected2()
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.NumberGroupSizes = NewLongs(2, 3)
    
    Actual = Convert.ToString(1, "n", Info)
    
    Assert.That Actual, Iz.EqualTo("1.00")
End Sub

Public Sub ToString_WithDoubleUsingFormatting_ReturnsExpected(ByVal Value As Double, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithSingleUsingFormatting_ReturnsExpected(ByVal Value As Single, ByVal Format As String, ByVal Expected As String)
    Dim Actual As String
    
    Actual = Convert.ToString(Value, Format)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithNumberNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As NumberNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.NumberNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "n", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCurrencyNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As CurrencyNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.CurrencyNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "c", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCurrencyPositivePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As CurrencyPositivePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.CurrencyPositivePattern = Pattern
    
    Actual = Convert.ToString(Value, "c", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithPercentNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As PercentNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Info.PercentNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "p", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomCurrencyGroupSizes_ReturnsCustomGrouping(ByVal Value As Double, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    Info.CurrencyGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "c", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomNumberGroupSizes_ReturnsCustomGrouping(ByVal Value As Double, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    Info.NumberGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "n", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomPercentageGroupSizes_ReturnsCustomGrouping(ByVal Value As Double, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim Info As NumberFormatInfo
    Set Info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    Info.PercentGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "p", Info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithNotANumber_ReturnsNAN()
    Dim d As Double
    On Error Resume Next
    d = 0# / 0#
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Convert.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("NaN")
End Sub

Public Sub ToString_WithPositiveInfinity_ReturnsInfinity()
    Dim d As Double
    On Error Resume Next
    d = 1# / 0#
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Convert.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("Infinity")
End Sub

Public Sub ToString_WithNegativeInfinity_ReturnMinusInfinity()
    Dim d As Double
    On Error Resume Next
    d = -1# / 0#
    On Error GoTo 0
    
    Dim Actual As String
    Actual = Convert.ToString(d)
    
    Assert.That Actual, Iz.EqualTo("-Infinity")
End Sub

Public Sub ToString_WithDoubleAndFormattingOfDecimal_Throws()
    On Error GoTo errTrap
    
    Convert.ToString 1#, "d"
    On Error GoTo 0
    Assert.Fail "An exception should be thrown."
    
errTrap:
    Dim Ex As Exception
    Assert.IsTrue Catch(Ex), "An exception should be caught."
    Assert.IsTrue TypeOf Ex Is FormatException, "Wrong exception type caught."
End Sub


Private Function MakeImaginary(ByVal Value As Double) As Double
    On Error Resume Next
    MakeImaginary = Value / 0#
    On Error GoTo 0
End Function


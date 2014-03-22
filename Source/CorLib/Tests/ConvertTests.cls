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
            Test.Use 12345678.91, NewLongs(1, 2, 3), "1,234,567,89,1.00 %"
            Test.Use 123456789.12, NewLongs(1, 2, 1, 0), "1234567,8,91,2.00 %"
            Test.Use 1234567890#, NewLongs(), "123456789000.00 %"
            Test.Use 1.234568E+07!, NewLongs(1, 2, 3), "1,234,568,00,0.00 %"
            
    End Select
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

Public Sub ToString_WithNumberNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As NumberNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.NumberNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "n", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCurrencyNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As CurrencyNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.CurrencyNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "c", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCurrencyPositivePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As CurrencyPositivePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.CurrencyPositivePattern = Pattern
    
    Actual = Convert.ToString(Value, "c", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithPercentNegativePattern_ReturnsExpected(ByVal Value As Variant, ByVal Pattern As PercentNegativePatterns, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    info.PercentNegativePattern = Pattern
    
    Actual = Convert.ToString(Value, "p", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomCurrencyGroupSizes_ReturnsCustomGrouping(ByVal Value As Double, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    info.CurrencyGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "c", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomNumberGroupSizes_ReturnsCustomGrouping(ByVal Value As Double, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    info.NumberGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "n", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ToString_WithCustomPercentageGroupSizes_ReturnsCustomGrouping(ByVal Value As Variant, ByRef GroupSizes As Variant, ByVal Expected As String)
    Dim Actual As String
    Dim info As NumberFormatInfo
    Set info = Cor.NewCultureInfo("en-US", False).NumberFormat
    Dim Groups() As Long
    Groups = GroupSizes
    info.PercentGroupSizes = Groups
    
    Actual = Convert.ToString(Value, "p", info)
    
    Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Private Function MakeImaginary(ByVal Value As Double) As Double
    On Error Resume Next
    MakeImaginary = Value / 0#
    On Error GoTo 0
End Function


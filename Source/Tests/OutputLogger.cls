VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutputLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IEventListener

Private Sub WriteResult(ByVal Result As TestResult)
    Dim FileNumber As Long
    Dim Message As String
    
    Message = Result.FullName & ": " & Result.Message
    FileNumber = FreeFile
    
    Open FileName For Append As #FileNumber
    Print #FileNumber, Message
    Close #FileNumber
End Sub

Private Property Get FileName() As String
    FileName = App.Path & "\TestOutput.txt"
End Property

Private Sub IEventListener_RunException(ByVal Info As SimplyVBComp.ErrorInfo)

End Sub

Private Sub IEventListener_RunFinished(ByVal Result As SimplyVBComp.TestResult)

End Sub

Private Sub IEventListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    If Len(Dir$(FileName)) > 0 Then
        Kill FileName
    End If
End Sub

Private Sub IEventListener_TestCaseFinished(ByVal Result As SimplyVBComp.TestResult)
    If Not Result.IsSuccess Then
        WriteResult Result
    End If
End Sub

Private Sub IEventListener_TestCaseStarted(ByVal Test As SimplyVBComp.TestCase)

End Sub

Private Sub IEventListener_TestOutput(ByVal Output As SimplyVBComp.TestOutput)

End Sub

Private Sub IEventListener_TestSuiteFinished(ByVal Result As SimplyVBComp.TestResult)

End Sub

Private Sub IEventListener_TestSuiteStarted(ByVal Suite As SimplyVBComp.TestSuite)

End Sub

Private Sub IEventListener_UnhandledError(ByVal Info As SimplyVBComp.ErrorInfo)

End Sub

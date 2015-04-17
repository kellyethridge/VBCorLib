VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
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
' Module: Console
'

''
' Provides a mechanism for handling input and output for either a screen or Stream.
'
' @remarks Using the <b>Console</b> class allows input and output to a console window.
' By default, a VB6 application will create a new console window for the <b>Console</b>
' class to use. Most times it is desired to run an application in an already open
' console window. If a VB6 application is run from within an existing console window,
' a new window will be created again. To have the VB6 application to use the console
' window from which is it run, the EXE needs to be converted to a console application.
' This can be accomplished by hand using the <i>Link</i> program that is included with
' Visual Basic 6. In a console window use the following command:
' <p>link /edit /subsystem:console <i>EXE Name</i></p>
' <p>Replace <EXE> with the path and executable to be flagged as a console application.
' Be sure the <i>Link</i> and application are both in paths the console window can find.
' Once the command has been executed, the VB6 application will run inside the console
' window from which it was run. If no window existed, then a new one will be created.</p>
' <p>The Console can be used to display text on the screen, and retrieve
' user input from a keyboard. The input and output can be redirected to a new
' source to allow for reading and writing from a specified Stream.</p>
' <p><b>WARNING</b>: Do not close the console through with <b>Close Window</b> button while
' working in the IDE. This will cause the IDE to crash and all data will be lost.</p>
'
' <pre>
' Console.BackgroundColor = Blue
' Console.WriteLine "hello"
' Console.ResetColor
' </pre>
'
' @see ConsoleKeyInfo
'
Option Explicit
Private Declare Function AllocConsole Lib "kernel32.dll" () As Long
Private Declare Function FreeConsole Lib "kernel32.dll" () As Long
Private Declare Function GetConsoleOutputCP Lib "kernel32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SetConsoleCtrlHandler Lib "kernel32.dll" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Private Declare Function BeepAPI Lib "kernel32.dll" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function SetConsoleWindowInfo Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, ByRef lpConsoleWindow As SMALL_RECT) As Long
Private Declare Function GetStdHandle Lib "kernel32.dll" (ByVal nStdHandle As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function SetConsoleScreenBufferSize Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef dwSize As Any) As Long
Private Declare Function CreateConsoleScreenBuffer Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwFlags As Long, ByRef lpScreenBufferData As Any) As Long
Private Declare Function SetConsoleActiveScreenBuffer Lib "kernel32.dll" (ByVal hConsoleOutput As Long) As Long
Private Declare Function SetConsoleTitleA Lib "kernel32.dll" (ByVal lpConsoleTitle As String) As Long
Private Declare Function SetConsoleTitleW Lib "kernel32.dll" (ByVal lpConsoleTitle As Long) As Long
Private Declare Function GetConsoleTitleA Lib "kernel32.dll" (ByVal lpConsoleTitle As String, ByVal nSize As Long) As Long
Private Declare Function GetConsoleTitleW Lib "kernel32.dll" (ByVal lpConsoleTitle As Long, ByVal nSize As Long) As Long
Private Declare Function GetLargestConsoleWindowSize Lib "kernel32.dll" (ByVal hConsoleOutput As Long) As COORD
Private Declare Function SetConsoleTextAttribute Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef dwCursorPosition As Any) As Long
Private Declare Function GetConsoleCursorInfo Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorInfo Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByRef lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function FillConsoleOutputCharacter Lib "kernel32.dll" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, ByRef dwWriteCoord As Any, ByRef lpNumberOfCharsWritten As Long) As Long
Private Declare Function FillConsoleOutputAttribute Lib "kernel32.dll" (ByVal hConsoleOutput As Long, ByVal wAttribute As Long, ByVal nLength As Long, ByRef dwWriteCoord As Any, ByRef lpNumberOfAttrsWritten As Long) As Long
Private Declare Function ReadConsoleOutput Lib "kernel32.dll" Alias "ReadConsoleOutputA" (ByVal hConsoleOutput As Long, ByRef lpBuffer As CHAR_INFO, ByRef dwBufferSize As Any, ByRef dwBufferCoord As Any, ByRef lpReadRegion As SMALL_RECT) As Long
Private Declare Function WriteConsoleOutput Lib "kernel32.dll" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, ByRef lpBuffer As CHAR_INFO, ByRef dwBufferSize As Any, ByRef dwBufferCoord As Any, ByRef lpWriteRegion As SMALL_RECT) As Long
Private Declare Function ReadConsoleInput Lib "kernel32.dll" Alias "ReadConsoleInputA" (ByVal hConsoleInput As Long, ByRef lpBuffer As INPUT_RECORD, ByVal nLength As Long, ByRef lpNumberOfEventsRead As Long) As Long
Private Declare Function SetConsoleOutputCP Lib "kernel32.dll" (ByVal wCodePageID As Long) As Long
Private Declare Function GetConsoleCP Lib "kernel32.dll" () As Long
Private Declare Function SetConsoleCP Lib "kernel32.dll" (ByVal wCodePageID As Long) As Long
Private Declare Function PeekConsoleInput Lib "kernel32.dll" Alias "PeekConsoleInputA" (ByVal hConsoleInput As Long, ByRef lpBuffer As INPUT_RECORD, ByVal nLength As Long, ByRef lpNumberOfEventsRead As Long) As Long
Private Declare Function GetConsoleMode Lib "kernel32.dll" (ByVal hConsoleHandle As Long, ByRef lpMode As Long) As Long
Private Declare Function SetConsoleMode Lib "kernel32.dll" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Private Const STD_OUTPUT_HANDLE         As Long = -11&
Private Const STD_INPUT_HANDLE          As Long = -10&
Private Const STD_ERROR_HANDLE          As Long = -12&
Private Const CONSOLE_TEXTMODE_BUFFER   As Long = 1
Private Const MaxTitleLength            As Long = 24500
Private Const MaxWindowSize             As Long = &H7FFF
Private Const KEY_EVENT                 As Long = &H1
Private Const NO_EVENT                  As Long = -1
Private Const ENABLE_PROCESSED_INPUT    As Long = &H1


Private Type CHAR_INFO
    Char        As Integer
    Attributes  As Integer
End Type

Private Type CONSOLE_CURSOR_INFO
    dwSize      As Long
    bVisible    As Long
End Type

Private Type COORD
    X As Integer
    Y As Integer
End Type

Private Type SMALL_RECT
    Left    As Integer
    Top     As Integer
    Right   As Integer
    Bottom  As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize              As COORD
    dwCursorPosition    As COORD
    wAttributes         As Integer
    srWindow            As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Type KEY_EVENT_RECORD
    bKeyDown            As Long
    wRepeatCount        As Integer
    wVirtualKeyCode     As Integer
    wVirtualScanCode    As Integer
    uChar               As Integer
    dwControlKeyState   As Long
End Type

Private Type INPUT_RECORD
    EventType   As Integer
    KeyEvent    As KEY_EVENT_RECORD
End Type


''
' A list of break types returned by the console.
'
' @param ControlC The CTRL+C key combination was pressed.
' @param ControlBreak The CTRL+Break key combination was pressed.
' @param CloseWindow The window is being closed.
' @param LogOffUser The user is logging off the system.
' @param ShutDownSystem The system is being shut down.
'
Public Enum ConsoleBreakType
    ControlC = 0
    ControlBreak = 1
    CloseWindow = 2
    LogOffUser = 5
    ShutDownSystem = 6
End Enum

''
' The set of colors used to foreground and background
' color settings of the console.
'
' @param CurrentColor Used to indicate to use the current color settings.
' @param Black
' @param DarkBlue
' @param DarkGreen
' @param DarkCyan
' @param DarkRed
' @param DarkMagenta
' @param DarkYellow
' @param Gray
' @param DarkGray
' @param Blue
' @param Green
' @param Cyan
' @param Red
' @param Magenta
' @param Yellow
' @param White
'
Public Enum ConsoleColor
    CurrentColor = -1
    Black = 0
    DarkBlue = 1
    DarkGreen = 2
    DarkCyan = 3
    DarkRed = 4
    DarkMagenta = 5
    DarkYellow = 6
    Gray = 7
    DarkGray = 8
    Blue = 9
    Green = 10
    Cyan = 11
    Red = 12
    Magenta = 13
    Yellow = 14
    White = 15
End Enum

Private mOwnsConsole        As Boolean
Private mOutput             As TextWriter
Private mInput              As TextReader
Private mError              As TextWriter
Private mOutputHandle       As Long
Private mInputHandle        As Long
Private mErrorHandle        As Long
Private mKeyPressHistory    As INPUT_RECORD
Private mOriginalOut        As Boolean
Private mOriginalError      As Boolean



''
' Returns the cause of the break event in the console
'
' @return The type of break that caused the interuption.
'
Public Property Get BreakType() As ConsoleBreakType
    BreakType = modConsoleCallbacks.BreakType
End Property

''
' Returns if a break has been caused.
'
' @return If a break has been caused in the console.
' @remarks This property will be set to true when one of
' several actions occurs. The actions are defined by the
' enumeration <b>ConsoleBreakType</b>. No DoEvents need
' to be called for this property to be set. Simply check
' this property when necessary to respond to any of the events.
'
Public Property Get Break() As Boolean
    Break = modConsoleCallbacks.Break
End Property

''
' Sets the break flag.
'
' @param RHS The break flag value.
' @remarks Set Break to false once it has been handled so the
' next console break can set it to true and be detected.
'
Public Property Let Break(ByVal RHS As Boolean)
    modConsoleCallbacks.Break = RHS
End Property

''
' Returns the Error output currently being used by the Console.
'
' @return An object that implements the TextWriter interface.
' @remarks The default Error writer writes to the Console error output.
'
Public Property Get Error() As TextWriter
    If mError Is Nothing Then
        Dim writer As StreamWriter
        Set writer = Cor.NewStreamWriter(OpenStandardError, Encoding.GetEncoding(GetConsoleOutputCP))
        writer.AutoFlush = True
        Set mError = writer
    End If
    Set Error = mError
End Property

''
' Returns the output writer currently being used by the Console.
'
' @return Returns an object that implements the TextWriter interface.
' @remarks The default output for the Console is to a Console window displayed
' on the screen. The output can be redirected to a new Stream using SetOut.
'
Public Property Get Out() As TextWriter
    If mOutput Is Nothing Then
        Dim writer As StreamWriter
        Set writer = Cor.NewStreamWriter(OpenStandardOutput, Encoding.GetEncoding(GetConsoleOutputCP))
        writer.AutoFlush = True
        Set mOutput = writer
    End If
    Set Out = mOutput
End Property

''
' Returns the input source currently used by the Console.
'
' @return Returns an object that implements the TextReader interface
' @remarks The default source of the input is the Keyboard. This can
' be changed by using the SetIn function.
'
Public Property Get In_() As TextReader
    If mInput Is Nothing Then Set mInput = Cor.NewStreamReader(OpenStandardInput, Encoding.GetEncoding(GetConsoleCP), False)
    Set In_ = mInput
End Property

''
' Reads a character from the current input source.
'
' @return The next character read from the input source, or -1 if no more characters exists.
' @remarks The Read method does not return until the Return key is pressed.
' Once the function returns, it can be called repeatedly until all key presses
' upto the return key are retrieved. The return key is also returned
' as a carriage-return (13) followed by a line-feed (10).
' @include "..\Includes\ConsoleRead.txt"
Public Function Read() As Long
    Read = In_.Read
End Function

''
' Reads all of the characters from the current input source until
' a NewLine break is reached.
'
' @return Returns all of the available characters upto a NewLine break,
' not including the NewLine characters.
' @remarks The ReadLine function will block and not return until a NewLine
' characters is pressed. This is usually the Return key.
' @include "..\Includes\ConsoleReadLine.txt"
Public Function ReadLine() As String
    ReadLine = In_.ReadLine
End Function

''
' Reads a single key from the console input.
'
' @param Intercept Flags whether the pressed key should be displayed
' through the console output or not.
' @return A <b>ConsoleKeyInfo</b> object containing information about the
' the key that was pressed and modifier key states.
' @remarks This method blocks the calling thread until a key is pressed.
' <p>Not all keys are read and returned. Modifier keys (Alt, Shift, Control) are
' not returned, but instead, included in the <b>ConsoleKeyInfo</b> object when
' a normal key is pressed. This allows the status of the modifier keys to be
' known during a keypress, such as a Shift-A key combination.</b>
' <p>Keys that do not trigger this method to return are ALT, SHIFT, CONTROL, CAPS-LOCK, SCROLL-LOCK, and NUM-LOCK.</p>
'
Public Function ReadKey(Optional ByVal Intercept As Boolean = False) As ConsoleKeyInfo
    If Not InternalReadKey(ReadKey) Then
        Do
            Dim Result  As Long
            Dim NumRead As Long
            Result = ReadConsoleInput(mInputHandle, mKeyPressHistory, 1, NumRead)
            If Result = 0 Or NumRead = 0 Then _
                Throw Cor.NewInvalidOperationException("Could not read from the console window.")
            
            Dim ReRead As Boolean
            ReRead = (mKeyPressHistory.EventType <> KEY_EVENT)
            If Not ReRead Then ReRead = (mKeyPressHistory.KeyEvent.bKeyDown = BOOL_FALSE)
            If Not ReRead And (mKeyPressHistory.KeyEvent.uChar = 0) Then
                Select Case mKeyPressHistory.KeyEvent.wVirtualKeyCode
                    Case vbKeyShift, vbKeyMenu, vbKeyControl, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital: ReRead = True
                End Select
            End If
        Loop While ReRead
        Call InternalReadKey(ReadKey)
    End If
    
    If Not Intercept Then Call WriteValue(ChrW$(mKeyPressHistory.KeyEvent.uChar))
End Function

''
' Sets the Console output to a new output TextWriter object.
'
' @param NewOut The TextWriter object that console output will be written to.
'
Public Sub SetOut(ByVal NewOut As TextWriter)
    If NewOut Is Nothing Then _
        Throw Cor.NewArgumentNullException("NewOut", "Cannot set output to Nothing.")
    
    Set mOutput = NewOut
    mOriginalOut = False
End Sub

''
' Sets the Console input source to a new TextReader object.
'
' @param NewIn The TextReader object the console retrieves input from.
'
Public Sub SetIn(ByVal NewIn As TextReader)
    If NewIn Is Nothing Then _
        Throw Cor.NewArgumentNullException("NewIn", "Cannot set input to Nothing.")
        
    Set mInput = NewIn
End Sub

''
' Sets the Console error output stream to the supplied TextWriter object.
'
' @param NewError The TextWriter object the console error data is written to.
'
Public Sub SetError(ByVal NewError As TextWriter)
    If NewError Is Nothing Then _
        Throw Cor.NewArgumentNullException("NewError", "Cannot set error output to Nothing.")
        
    Set mError = NewError
    mOriginalError = False
End Sub

''
' Returns the standard output stream.
'
' @return The standard output stream to write user data to.
' @remarks The standard output is the screen console.
'
Public Function OpenStandardOutput() As Stream
    Set OpenStandardOutput = GetConsoleStream(mOutputHandle, WriteAccess)
End Function

''
' Returns the standard input stream.
'
' @return The standard input stream object to retrieve user input from.
' @remarks The standard input is from the keyboard console.
'
Public Function OpenStandardInput() As Stream
    Set OpenStandardInput = GetConsoleStream(mInputHandle, ReadAccess)
End Function

''
' Returns the standard error stream.
'
' @return The standard error output stream that error data is written to.
' @remarks The standard error output is to the screen console.
'
Public Function OpenStandardError() As Stream
    Set OpenStandardError = GetConsoleStream(mErrorHandle, WriteAccess)
End Function

''
' Writes a value to the current output stream.
'
' @param values A set of zero or more parameters used to create a formatted string output.
'
' @remarks The first parameter is the string that will be written to
' the current output stream. All additional parameters will be used as
' formatting information in the first parameter. If no parameters are
' provided, then an empty string is written with a NewLine.
' @include "..\Includes\ConsoleWriteLine.txt"
Public Sub WriteLine(ParamArray Values() As Variant)
    Dim a() As Variant
    Call Helper.Swap4(ByVal ArrPtr(a), ByVal Helper.DerefEBP(12))
    Call InternalWriteLine(a)
End Sub

''
' Writes a string to the Console output stream without a NewLine break.
'
' @param value The value to write the string version of to the output stream.
' @param Args Values to be used in a formatted output.
' @remarks Values are converted to their string representation for text output.
' A formatted output can be created similar to <b>cString.Format</b>.
' @include "..\Includes\ConsoleWriteValue.txt"
Public Sub WriteValue(ByRef Value As Variant, ParamArray Args() As Variant)
    Dim a() As Variant
    Call Helper.Swap4(ByVal ArrPtr(a), ByVal Helper.DerefEBP(16))
    Call InternalWriteValue(Value, a)
End Sub

''
' Produces a beep through the console.
'
' @param Frequency The frequency the beep will play at (default 800hz).
' @param Duration The milliseconds the beep should play for (default 200ms).
'
Public Sub Beep(Optional ByVal Frequency As Long = 800, Optional ByVal Duration As Long = 200)
    Const MinFrequency  As Long = 37
    Const MaxFrequency  As Long = 32767
    Const MinDuration   As Long = 1
    
    If Frequency < MinFrequency Or Frequency > MaxFrequency Then _
        Throw Cor.NewArgumentOutOfRangeException("Frequency", Frequency, Environment.GetResourceString(ArgumentOutOfRange_BeepFrequency, MinFrequency, MaxFrequency))
    If Duration < MinDuration Then _
        Throw Cor.NewArgumentOutOfRangeException("Duration", Message:=Environment.GetResourceString(ArgumentOutOfRange_NeedPosNum))
    
    If BeepAPI(Frequency, Duration) = BOOL_FALSE Then Throw Cor.NewSystemException(GetSystemMessage(Err.LastDllError))
End Sub

''
' Sets the size of the viewable window area for the console buffer.
'
' @param Width The width in characters.
' @param Height The height in characters.
' @param The underlying buffer is expanded to accomodate a larger viewing window.
' The buffer is not reduced in size for smaller viewing windows.
' @remarks The screen buffer is the virtual screen that is written to. The size
' of the buffer can be larger than the window inwhich it is being viewed through.
' The window's size can be changed to show a different size of the underlying
' screen buffer without altering the size of that buffer.
'
Public Sub SetWindowSize(ByVal Width As Long, ByVal Height As Long)
    If Width <= 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Width", Width, Environment.GetResourceString(ArgumentOutOfRange_NeedPosNum))
    If Height <= 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Height", Height, Environment.GetResourceString(ArgumentOutOfRange_NeedPosNum))
    
    With GetBufferInfo
        Dim ExpandBuffer    As Boolean
        Dim NewBufferSize   As COORD
        
        NewBufferSize = .dwSize
        Select Case .srWindow.Left + Width
            Case Is > MaxWindowSize
                Throw Cor.NewArgumentOutOfRangeException("Width", Width, Environment.GetResourceString(ArgumentOutOfRange_ConsoleBufferSize))
                
            Case Is > .dwSize.X
                NewBufferSize.X = .srWindow.Left + Width
                ExpandBuffer = True
        End Select
        Select Case .srWindow.Top + Height
            Case Is > MaxWindowSize
                Throw Cor.NewArgumentOutOfRangeException("Height", Height, Environment.GetResourceString(ArgumentOutOfRange_ConsoleBufferSize))
                
            Case Is > .dwSize.Y
                NewBufferSize.Y = .srWindow.Top + Height
                ExpandBuffer = True
        End Select
        
        If ExpandBuffer Then
            If SetConsoleScreenBufferSize(mOutputHandle, ByVal AsLong(NewBufferSize)) = BOOL_FALSE Then IOError Err.LastDllError
        End If
        
        Dim NewWindowSize As SMALL_RECT
        NewWindowSize = .srWindow
        NewWindowSize.Right = .srWindow.Left + Width - 1
        NewWindowSize.Bottom = .srWindow.Top + Height - 1
        If SetConsoleWindowInfo(mOutputHandle, True, NewWindowSize) = BOOL_FALSE Then
            Dim e As Long
            e = Err.LastDllError
            
            If ExpandBuffer Then
                ' restore buffer to original size
                If SetConsoleScreenBufferSize(mOutputHandle, ByVal AsLong(.dwSize)) = BOOL_FALSE Then IOError Err.LastDllError
            End If
            
            Dim WindowSize As COORD
            WindowSize = GetLargestConsoleWindowSize(mOutputHandle)
            If Width > WindowSize.X Then _
                Throw Cor.NewArgumentOutOfRangeException("Width", WindowSize.X, Environment.GetResourceString(ArgumentOutOfRange_ConsoleWindowSize_Size, WindowSize.X))
            If Height > WindowSize.Y Then _
                Throw Cor.NewArgumentOutOfRangeException("Height", WindowSize.Y, Environment.GetResourceString(ArgumentOutOfRange_ConsoleWindowSize_Size, WindowSize.Y))
            
            IOError e
        End If
    End With
End Sub

''
' Sets the position of the viewing window within the console screen buffer.
'
' @param Left The left position in characters.
' @param Top The top position in characters.
' @remarks The underlying screen buffer is not expanded to accomodate moving
' the window outside of the screen buffer.
'
Public Sub SetWindowPosition(ByVal Left As Long, ByVal Top As Long)
    If Left < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Left", Left, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    If Top < 0 Then _
        Throw Cor.NewArgumentOutOfRangeException("Top", Top, Environment.GetResourceString(ArgumentOutOfRange_NeedNonNegNum))
    
    With GetBufferInfo
        If Left + (.srWindow.Right - .srWindow.Left + 1) > .dwSize.X Then _
            Throw Cor.NewArgumentOutOfRangeException("Left", Left, Environment.GetResourceString(ArgumentOutOfRange_ConsoleWindowPos))
        If Top + (.srWindow.Bottom - .srWindow.Top + 1) > .dwSize.Y Then _
            Throw Cor.NewArgumentOutOfRangeException("Top", Top, Environment.GetResourceString(ArgumentOutOfRange_ConsoleWindowPos))
        
        Dim NewPosition As SMALL_RECT
        NewPosition.Left = Left
        NewPosition.Top = Top
        With .srWindow
            NewPosition.Right = Left + (.Right - .Left)
            NewPosition.Bottom = Top + (.Bottom - .Top)
        End With
        If SetConsoleWindowInfo(mOutputHandle, True, NewPosition) = BOOL_FALSE Then IOError Err.LastDllError
    End With
End Sub

''
' Sets the size of the underlying screen buffer.
'
' @param Width The width in characters of the buffer.
' @param height The height in characters of the buffer.
'
Public Sub SetBufferSize(ByVal Width As Long, ByVal Height As Long)
    If Width < 1 Or Width > MaxWindowSize Then _
        Throw Cor.NewArgumentOutOfRangeException("Width must be between 1 and 32767.", "Width", Width)
    If Height < 1 Or Height > MaxWindowSize Then _
        Throw Cor.NewArgumentOutOfRangeException("Height must be between 1 and 32767.", "Height", Height)
    
    With GetBufferInfo.srWindow
        If Width < .Right - .Left + 1 Then _
            Throw Cor.NewArgumentOutOfRangeException("Width", Width, Environment.GetResourceString(ArgumentOutOfRange_ConsoleBufferLessThanWindowSize))
        If Height < .Bottom - .Top + 1 Then _
            Throw Cor.NewArgumentOutOfRangeException("Height", Height, Environment.GetResourceString(ArgumentOutOfRange_ConsoleBufferLessThanWindowSize))
    End With
    
    Dim NewSize As COORD
    With NewSize
        .X = Width
        .Y = Height
    End With
    If SetConsoleScreenBufferSize(mOutputHandle, ByVal AsLong(NewSize)) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

''
' Returns the current width of the underlying screen buffer in characters.
'
' @return The width of the screen buffer.
'
Public Property Get BufferWidth() As Long
    BufferWidth = GetBufferInfo.dwSize.X
End Property

''
' Sets the width of the underlying screen buffer in characters.
'
' @param RHS The new width of the screen buffer.
'
Public Property Let BufferWidth(ByVal RHS As Long)
    Call SetBufferSize(RHS, BufferHeight)
End Property

''
' Returns the current height of the underlying screen buffer in characters.
'
' @return The height of the screen buffer.
'
Public Property Get BufferHeight() As Long
    BufferHeight = GetBufferInfo.dwSize.Y
End Property

''
' Sets the height of the underlying screen buffer in characters.
'
' @param RHS The new height of the screen buffer.
'
Public Property Let BufferHeight(ByVal RHS As Long)
    Call SetBufferSize(BufferWidth, RHS)
End Property

''
' Returns the left coordinate of the console window within the screen buffer
' in character columns.
'
' @return The console window left coordinate in character columns.
' @remarks The underlying screen buffer can be larger than the console
' window itself. In order to see other portions of the screen buffer, the
' window must scroll around the buffer. If the screen buffer is wider than
' the window itself, then the window must be scrolled horizontally to view
' the screen buffer.
'
Public Property Get WindowLeft() As Long
    WindowLeft = GetBufferInfo.srWindow.Left
End Property

''
' Sets the left coordinate of the console window within the screen buffer
' in character columns.
'
' @param RHS The new left position.
' @remarks The view window maitains the original size.
'
Public Property Let WindowLeft(ByVal RHS As Long)
    Call SetWindowPosition(RHS, WindowTop)
End Property

''
' Returns the top of the viewable window within the screen buffer.
'
' @return The top coordinate of the window.
'
Public Property Get WindowTop() As Long
    WindowTop = GetBufferInfo.srWindow.Top
End Property

''
' Sets the top of the viewable windwo within the screen buffer.
'
' @param RHS The new top position in characters, starting with 0.
'
Public Property Let WindowTop(ByVal RHS As Long)
    Call SetWindowPosition(WindowLeft, RHS)
End Property

''
' Returns the width of the window in characters.
'
' @return The width of the viewable window.
'
Public Property Get WindowWidth() As Long
    With GetBufferInfo.srWindow
        WindowWidth = .Right - .Left + 1
    End With
End Property

''
' Sets the width of the viewable window within the screen buffer.
'
' @param RHS The new width of the window in characters.
'
Public Property Let WindowWidth(ByVal RHS As Long)
    Call SetWindowSize(RHS, WindowHeight)
End Property

''
' Returns the height of the viewable window in the screen buffer.
'
' @return The height of the window in characters.
'
Public Property Get WindowHeight() As Long
    With GetBufferInfo.srWindow
        WindowHeight = .Bottom - .Top + 1
    End With
End Property

''
' Sets the height of the viewable window in the screen buffer.
'
' @param RHS The new window height in characters.
'
Public Property Let WindowHeight(ByVal RHS As Long)
    Call SetWindowSize(WindowWidth, RHS)
End Property

''
' Returns the largest width the viewable window could be set to.
'
' @return The largest possible window width.
'
Public Property Get LargestWindowWidth() As Long
    LargestWindowWidth = GetLargestConsoleWindowSize(mOutputHandle).X
End Property

''
' Returns the largest height the viewable window could be set to.
'
' @return The largest possible window height.
'
Public Property Get largestWindowHeight() As Long
    largestWindowHeight = GetLargestConsoleWindowSize(mOutputHandle).Y
End Property

''
' Returns the console title of the current process.
'
' @return The console title.
' @remarks This method support Unicode characters on NT platforms.
'
Public Property Get Title() As String
    Dim Buf As String
    Dim Size As Long
    
    Buf = String$(MaxTitleLength, 0)
    If Environment.IsNT Then
        Size = GetConsoleTitleW(StrPtr(Buf), MaxTitleLength)
    Else
        Size = GetConsoleTitleA(Buf, MaxTitleLength)
    End If
    Select Case Size
        Case 0
            Dim e As Long
            e = Err.LastDllError
            If e <> 0 Then IOError e
        Case Is > MaxTitleLength
            Throw Cor.NewInvalidOperationException("Cannot retrieve title longer that 24500 characters.")
    End Select
    Title = Left$(Buf, Size)
End Property

''
' Sets the title for the console of the current process.
'
' @param RHS The title to set the console to.
' @remarks This method supports Unicode characters on NT platforms.
'
Public Property Let Title(ByVal RHS As String)
    If Len(RHS) > MaxTitleLength Then _
        Throw Cor.NewArgumentOutOfRangeException("Title", RHS, Environment.GetResourceString(ArgumentOutOfRange_ConsoleTitleTooLong))
    
    If cString.IsNull(RHS) Then RHS = ""
    
    If Environment.IsNT Then
        If SetConsoleTitleW(StrPtr(RHS)) = BOOL_FALSE Then IOError Err.LastDllError
    Else
        If SetConsoleTitleA(RHS) = BOOL_FALSE Then IOError Err.LastDllError
    End If
End Property

''
' Returns the current background color.
'
' @return The current background color for text.
' @remarks The background color is painted individually for each
' text characters displayed. This is not the same as the entire
' window background color.
'
Public Property Get BackgroundColor() As ConsoleColor
    BackgroundColor = (GetBufferInfo.wAttributes And &HF0) \ &H10
End Property

''
' Sets the background color.
'
' @param RHS The new background color.
' @remarks The new background color will only affect text displayed
' after the changing of the background color.
'
Public Property Let BackgroundColor(ByVal RHS As ConsoleColor)
    If RHS < 0 Or RHS > 15 Then _
        Throw Cor.NewArgumentOutOfRangeException("BackgroundColor", RHS, Environment.GetResourceString(ArgumentOutOfRange_ConsoleColor))
    
    If SetConsoleTextAttribute(mOutputHandle, (GetBufferInfo.wAttributes And &HFF0F) Or (RHS * &H10)) = BOOL_FALSE Then IOError Err.LastDllError
End Property

''
' Returns the current foreground color.
'
' @return The current foreground color for text.
' @remarks The foreground color is painted individually for each
' text characters displayed.
'
Public Property Get ForegroundColor() As ConsoleColor
    ForegroundColor = GetBufferInfo.wAttributes And &HF
End Property

''
' Sets the foreground color.
'
' @param RHS The new foreground color.
' @remarks The new foreground color will only affect text that is
' displayed after the color has been changed.
'
Public Property Let ForegroundColor(ByVal RHS As ConsoleColor)
    If RHS < 0 Or RHS > 15 Then _
        Throw Cor.NewArgumentOutOfRangeException("ForegroundColor", RHS, Environment.GetResourceString(ArgumentOutOfRange_ConsoleColor))
    
    If SetConsoleTextAttribute(mOutputHandle, (GetBufferInfo.wAttributes And &HFFF0) Or RHS) = BOOL_FALSE Then IOError Err.LastDllError
End Property

''
' Sets the position of the cursor within the screen buffer.
'
' @param Left The number of characters from the left of the screen buffer.
' @param Top The number of characters from the top of the screen buffer.
' @remarks The positions start at 0,0 and go to BufferWidth - 1, BufferHeight - 1.
'
Public Sub SetCursorPosition(ByVal Left As Long, ByVal Top As Long)
    With GetBufferInfo.dwSize
        If Left < 0 Or Left >= .X Then _
            Throw Cor.NewArgumentOutOfRangeException("Left", Left, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If Top < 0 Or Top >= .Y Then _
            Throw Cor.NewArgumentOutOfRangeException("Top", Top, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
    End With
    
    Dim NewPosition As COORD
    With NewPosition
        .X = Left
        .Y = Top
    End With
    If SetConsoleCursorPosition(mOutputHandle, ByVal AsLong(NewPosition)) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

''
' Returns the left coordinate of the cursor position.
'
' @return The left coordinate.
'
Public Property Get CursorLeft() As Long
    CursorLeft = GetBufferInfo.dwCursorPosition.X
End Property

''
' Sets the left coordinate of the cursor position.
'
' @param RHS The new left coordinate.
'
Public Property Let CursorLeft(ByVal RHS As Long)
    Call SetCursorPosition(RHS, CursorTop)
End Property

''
' Returns the top coordinate of the cursor position.
'
' @return The top coordinate.
'
Public Property Get CursorTop() As Long
    CursorTop = GetBufferInfo.dwCursorPosition.Y
End Property

''
' Sets the top coordinate for the cursor position
'
' @param RHS The top coordinate.
'
Public Property Let CursorTop(ByVal RHS As Long)
    Call SetCursorPosition(CursorLeft, RHS)
End Property

''
' Returns the size (height) of the cursor.
'
' @return The size of the cursor.
' @remarks the size of the cursor is a value 1 to 100, 1 being
' 1% the full block, and 100 beign 100% a full block.
'
Public Property Get CursorSize() As Long
    CursorSize = GetCursorInfo.dwSize
End Property

''
' Sets the size (height) of the cursor.
'
' @param RHS The new cursor size.
' @remarks the size of the cursor is a value 1 to 100, 1 being
' 1% the full block, and 100 being 100% a full block.
'
Public Property Let CursorSize(ByVal RHS As Long)
    If RHS < 1 Or RHS > 100 Then _
        Throw Cor.NewArgumentOutOfRangeException("CursorSize", RHS, Environment.GetResourceString(ArgumentOutOfRange_CursorSize))
    
    Dim Info As CONSOLE_CURSOR_INFO
    With Info
        .bVisible = GetCursorInfo.bVisible
        .dwSize = RHS
    End With
    If SetConsoleCursorInfo(mOutputHandle, Info) = BOOL_FALSE Then IOError Err.LastDllError
End Property

''
' Returns whether or not the cursor is currently invisible.
'
' @returns The visibility of the cursor.
'
Public Property Get CursorVisible() As Boolean
    CursorVisible = GetCursorInfo.bVisible
End Property

''
' Sets if the cursor is visible or not.
'
' @param RHS The visibility of the cursor.
'
Public Property Let CursorVisible(ByVal RHS As Boolean)
    Dim Info As CONSOLE_CURSOR_INFO
    
    With Info
        .bVisible = RHS
        .dwSize = CursorSize
    End With
    If SetConsoleCursorInfo(mOutputHandle, Info) = BOOL_FALSE Then IOError Err.LastDllError
End Property

''
' Clears the console screen and window.
'
Public Sub Clear()
    Dim CharsToWrite As Long
    With GetBufferInfo.dwSize
        CharsToWrite = .X * .Y
    End With
    
    Dim StartPosition As COORD
    If FillConsoleOutputCharacter(mOutputHandle, 32, CharsToWrite, ByVal AsLong(StartPosition), 0) = BOOL_FALSE Then IOError Err.LastDllError
    If FillConsoleOutputAttribute(mOutputHandle, GetBufferInfo.wAttributes, CharsToWrite, ByVal AsLong(StartPosition), 0) = BOOL_FALSE Then IOError Err.LastDllError
    If SetConsoleCursorPosition(mOutputHandle, ByVal AsLong(StartPosition)) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

''
' Moves one section of the screen buffer to another location within the screen buffer.
'
' @param SourceLeft The left side of the area to be moved.
' @param SourceTop The top side of the area to be moved.
' @param SourceWidth The width of the area to be moved.
' @param SourceHeight The height of the area to be moved.
' @param TargetLeft The starting left position to move the area to.
' @param TargetTop The starting top position to move the area to.
' @param SourceChar The character to fill the moved area in with. This can be a String or Ascii value.
' @param SourceForeColor The forecolor used when filling in the original moved area.
' @param SourceBackColor The backcolor used when filling in the original moved area.
'
Public Sub MoveBufferArea(ByVal SourceLeft As Long, ByVal SourceTop As Long, ByVal SourceWidth As Long, ByVal SourceHeight As Long, ByVal TargetLeft As Long, ByVal TargetTop As Long, Optional ByVal SourceChar As Variant = " ", Optional ByVal SourceForeColor As ConsoleColor = ConsoleColor.Black, Optional ByVal SourceBackColor As ConsoleColor = CurrentColor)
    Dim Buffer() As CHAR_INFO
    Dim BufferSize As COORD
    Dim Region As SMALL_RECT
    
    With GetBufferInfo.dwSize
        If SourceLeft < 0 Or SourceLeft >= .X Then _
            Throw Cor.NewArgumentOutOfRangeException("SourceLeft", SourceLeft, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If SourceTop < 0 Or SourceTop >= .Y Then _
            Throw Cor.NewArgumentOutOfRangeException("SourceTop", SourceTop, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If SourceWidth < 0 Or SourceWidth > (.X - SourceLeft) Then _
            Throw Cor.NewArgumentOutOfRangeException("SourceWidth", SourceWidth, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If SourceWidth < 0 Or SourceWidth > (.Y - SourceTop) Then _
            Throw Cor.NewArgumentOutOfRangeException("SourceHeight", SourceHeight, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If TargetLeft < 0 Or TargetLeft >= .X Then _
            Throw Cor.NewArgumentOutOfRangeException("TargetLeft", TargetLeft, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If TargetTop < 0 Or TargetTop >= .Y Then _
            Throw Cor.NewArgumentOutOfRangeException("TargetTop", TargetTop, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
    End With
        
    If SourceWidth = 0 Or SourceHeight = 0 Then Exit Sub
    
    ReDim Buffer(SourceWidth * SourceHeight - 1)
    With BufferSize
        .X = SourceWidth
        .Y = SourceHeight
    End With
    With Region
        .Left = SourceLeft
        .Right = SourceLeft + SourceWidth - 1
        .Top = SourceTop
        .Bottom = SourceTop + SourceHeight - 1
    End With
    If ReadConsoleOutput(mOutputHandle, Buffer(0), ByVal AsLong(BufferSize), ByVal 0&, Region) = BOOL_FALSE Then IOError Err.LastDllError
    
    Call InternalFillArea(SourceLeft, SourceTop, SourceWidth, SourceHeight, SourceChar, SourceForeColor, SourceBackColor)
    
    With Region
        .Left = TargetLeft
        .Right = TargetLeft + SourceWidth - 1
        .Top = TargetTop
        .Bottom = TargetTop + SourceHeight - 1
    End With
    If WriteConsoleOutput(mOutputHandle, Buffer(0), ByVal AsLong(BufferSize), ByVal 0&, Region) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

''
' Fills a specified region in the screen buffer with a character and color attributes.
'
' @param Left The left position in the screen buffer to start filling from.
' @param Top The top position in the screen buffer to start filling from.
' @param Width The width of the area to be filled.
' @param Height The height of the area to be filled.
' @param FillChar The character to fill the area with.
' @param ForeColor The color of the fill character.
' @param BackColor The color of the background behind the fill character.
'
Public Sub FillBufferArea(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal FillChar As Variant = " ", Optional ByVal Forecolor As ConsoleColor = ConsoleColor.CurrentColor, Optional ByVal BackColor As ConsoleColor = ConsoleColor.CurrentColor)
    With GetBufferInfo.dwSize
        If Left < 0 Or Left >= .X Then _
            Throw Cor.NewArgumentOutOfRangeException("Left", Left, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If Top < 0 Or Top >= .Y Then _
            Throw Cor.NewArgumentOutOfRangeException("Top", Top, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If Width < 0 Or Width > (.X - Left) Then _
            Throw Cor.NewArgumentOutOfRangeException("Width", Width, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
        If Width < 0 Or Width > (.Y - Top) Then _
            Throw Cor.NewArgumentOutOfRangeException("Height", Height, Environment.GetResourceString(ArgumentOutOfRange_OutsideConsoleBoundry))
    End With

    Call InternalFillArea(Left, Top, Width, Height, FillChar, Forecolor, BackColor)
End Sub

''
' Clears a specificed region in the screen buffer.
'
' @param Left The left position in the screen buffer to start clearing from.
' @param Top The top position in the screen buffer to start clearing from.
' @param Width The width of the area to be cleared.
' @param Height The height of the area to be cleared.
' @param ForeColor The color to set the text foreground color to in the cleared area.
' @param BackColor The color to set the text background color to in the cleared area.
'
Public Sub ClearBufferArea(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal Forecolor As ConsoleColor = ConsoleColor.CurrentColor, Optional ByVal BackColor As ConsoleColor = ConsoleColor.CurrentColor)
    Call FillBufferArea(Left, Top, Width, Height, " ", Forecolor, BackColor)
End Sub

''
' Restores the foreground and background colors to their original values.
'
' @remarks The backcolor will be set to ConsoleColor.Black, the forecolor
' will be set to ConsoleColor.Gray.
'
Public Sub ResetColor()
    If SetConsoleTextAttribute(mOutputHandle, (GetBufferInfo.wAttributes And &HFF00) Or ConsoleColor.Gray) = BOOL_FALSE Then IOError Err.LastDllError
End Sub

''
' Returns the encoding for the current console output.
'
' @return An Encoding object for the current console output encoding.
'
Public Property Get OutputEncoding() As Encoding
    Set OutputEncoding = Encoding.GetEncoding(GetConsoleOutputCP)
End Property

''
' Sets the encoding for the console output.
'
' @param RHS The new encoding used for output to the console.
' @remarks This is only supported on Windows NT machines.
'
Public Property Set OutputEncoding(ByVal RHS As Encoding)
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("OutputEncoding", "Cannot set output encoding to Nothing.")
        
    Call Environment.VerifyNTMachine
    
    If (Not mOutput Is Nothing) And (mOriginalOut = True) Then
        mOutput.Flush
        Set mOutput = Nothing
    End If
    If (Not mError Is Nothing) And (mOriginalError = True) Then
        mError.Flush
        Set mError = Nothing
    End If
    
    If SetConsoleOutputCP(RHS.CodePage) = BOOL_FALSE Then IOError Err.LastDllError
End Property

''
' Returns the encoding used for the input from the console.
'
' @return An encoding for the input from the console.
'
Public Property Get InputEncoding() As Encoding
    Set InputEncoding = Encoding.GetEncoding(GetConsoleCP)
End Property

''
' Sets the encoding used during input from the console.
'
' @param RHS The new encoding used for input from the console.
' @remarks This is only supported on Windows NT machines.
'
Public Property Set InputEncoding(ByVal RHS As Encoding)
    If RHS Is Nothing Then _
        Throw Cor.NewArgumentNullException("InputEncoding", "Cannot set input encoding to Nothing.")
        
    Call Environment.VerifyNTMachine
    
    If SetConsoleCP(RHS.CodePage) = BOOL_FALSE Then IOError Err.LastDllError
    Set mInput = Nothing
End Property

''
' Returns the state of the Caps-Lock key.
'
' @return CapsLock On = True, otherwise False.
' @remarks The CapsLock key is a toggle key, so this property
' returns the current state, not if it is currently being pressed
' or not, only if it is on or off.
'
Public Property Get CapsLock() As Boolean
    CapsLock = (GetKeyState(vbKeyCapital) And 1)
End Property

''
' Returns the state of the Num-Lock key.
' @return NumLock On = True, otherwise False.
' @remarks The NumLock key is a toggle key, so this property
' returns the current state, not if it is currently being pressed
' or not, only if it is on or off.
'
Public Property Get NumLock() As Boolean
    NumLock = (GetKeyState(vbKeyNumlock) And 1)
End Property

''
' Returns if a key has been pressed and is available to be
' read in using one of the read methods.
'
' @return Returns if a key is ready to be read in.
'
Public Property Get KeyAvailable() As Boolean
    Dim Record  As INPUT_RECORD
    Dim NumRead As Long
    
    If mKeyPressHistory.EventType = KEY_EVENT Then
        KeyAvailable = True
        Exit Property
    End If
    
    Do
        If PeekConsoleInput(mInputHandle, Record, 1, NumRead) = BOOL_FALSE Then IOError Err.LastDllError
        If NumRead = 0 Then Exit Property
        If Record.EventType = KEY_EVENT And Record.KeyEvent.bKeyDown = BOOL_TRUE Then
            KeyAvailable = True
            Exit Property
        End If
        
        ' nothing matched what we expected, so read it and discard.
        If ReadConsoleInput(mInputHandle, Record, 1, NumRead) = BOOL_FALSE Then IOError Err.LastDllError
    Loop
End Property

''
' Returns if the Control+C combination is managed by the system or returned as keyboard input.
'
' @return The state of management for Control+C.
'
Public Property Get TreatControlCAsInput() As Boolean
    Dim modes As Long
    If GetConsoleMode(mInputHandle, modes) = BOOL_FALSE Then IOError Err.LastDllError
    TreatControlCAsInput = (modes And ENABLE_PROCESSED_INPUT)
End Property

''
' Sets if the Control+C combination is managed by the system or returned as keyboard input.
'
' @param RHS Flag indicating if the system should manage Control+C or not.
'
Public Property Let TreatControlCAsInput(ByVal RHS As Boolean)
    Dim modes As Long
    If GetConsoleMode(mInputHandle, modes) = BOOL_FALSE Then IOError Err.LastDllError
    If RHS Then
        modes = modes Or ENABLE_PROCESSED_INPUT
    Else
        modes = modes And (Not ENABLE_PROCESSED_INPUT)
    End If
    If SetConsoleMode(mInputHandle, modes) = BOOL_FALSE Then IOError Err.LastDllError
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetConsoleStream(ByVal Handle As Long, ByVal Access As FileAccess) As Stream
    Dim Ret As New ConsoleStream
    Call Ret.Init(Handle, Access)
    Set GetConsoleStream = Ret
End Function

Private Sub InternalWrite(ByVal Value As String, ByRef Args() As Variant, ByVal NewLine As Boolean)
    If Not cArray.IsNull(Args) Then
        If UBound(Args) >= 0 Then Value = cString.FormatArray(Value, Args)
    End If
    If NewLine Then Value = Value & Out.NewLine
    Call Out.WriteValue(Value)
End Sub

Private Sub InternalWriteLine(ByRef Args() As Variant)
    Dim ub  As Long
    Dim s   As String
    
    ub = UBound(Args)
    
    If ub >= 0 Then s = Object.ToString(Args(0))
    If ub > 0 Then
        Dim i As Long
        For i = 1 To ub
            Call Helper.MoveVariant(Args(i - 1), Args(i))
        Next i
    Else
        Erase Args
    End If
    
    Call InternalWrite(s, Args, True)
End Sub

Private Sub InternalWriteValue(ByRef Value As Variant, ByRef Args() As Variant)
    InternalWrite Convert.ToString(Value), Args, False
End Sub

Private Sub Init()
    mOwnsConsole = (AllocConsole <> BOOL_FALSE)
    Call Sleep(10)     ' give it time to be created
    
    If InDebugger Then
        Call DisableCloseButton
    End If
    
    Call SetConsoleCtrlHandler(AddressOf ControlBreakHandler, True)
    mOutputHandle = GetStdHandle(STD_OUTPUT_HANDLE)
    mInputHandle = GetStdHandle(STD_INPUT_HANDLE)
    mErrorHandle = GetStdHandle(STD_ERROR_HANDLE)
    mKeyPressHistory.EventType = NO_EVENT
    mOriginalOut = True
    mOriginalError = True
End Sub

Private Sub DisableCloseButton()
    Dim oldTitle As String
    oldTitle = Me.Title
    
    Dim pid As Long
    pid = GetCurrentProcessId()
    
    Dim newTitle As String
    newTitle = oldTitle & pid
    
    Me.Title = newTitle
    Call Sleep(40)
    
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, newTitle)
    
    Me.Title = oldTitle
    
    If hwnd <> vbNullPtr Then
        Dim hMenu As Long
        hMenu = GetSystemMenu(hwnd, False)
        Call RemoveMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
    End If
End Sub

Private Function GetBufferInfo() As CONSOLE_SCREEN_BUFFER_INFO
    If GetConsoleScreenBufferInfo(mOutputHandle, GetBufferInfo) = BOOL_FALSE Then IOError Err.LastDllError
End Function

Private Function GetCursorInfo() As CONSOLE_CURSOR_INFO
    If GetConsoleCursorInfo(mOutputHandle, GetCursorInfo) = BOOL_FALSE Then IOError Err.LastDllError
End Function

Private Function InternalReadKey(ByRef RetVal As ConsoleKeyInfo) As Boolean
    Const ALT_KEYS As Long = 3
    Const SHIFT_KEYS As Long = 16
    Const CTRL_KEYS As Long = 12
    
    If mKeyPressHistory.EventType <> KEY_EVENT Then Exit Function
    
    With mKeyPressHistory.KeyEvent
        .wRepeatCount = .wRepeatCount - 1
        If .wRepeatCount = 0 Then mKeyPressHistory.EventType = NO_EVENT
        Set RetVal = Cor.NewConsoleKeyInfo(.uChar, .wVirtualKeyCode, .dwControlKeyState And SHIFT_KEYS, .dwControlKeyState And ALT_KEYS, .dwControlKeyState And CTRL_KEYS)
    End With
    
    InternalReadKey = True
End Function

Private Sub InternalFillArea(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByRef FillChar As Variant, ByVal Forecolor As ConsoleColor, ByVal BackColor As ConsoleColor)
    
    Select Case Forecolor
        Case ConsoleColor.CurrentColor: Forecolor = ForegroundColor
        Case Is < 0, Is > 15
            Throw Cor.NewArgumentOutOfRangeException("ForeColor", Forecolor, Environment.GetResourceString(ArgumentOutOfRange_ConsoleColor))
    End Select

    Select Case BackColor
        Case ConsoleColor.CurrentColor: BackColor = BackgroundColor
        Case Is < 0, Is > 15
            Throw Cor.NewArgumentOutOfRangeException("BackColor", BackColor, Environment.GetResourceString(ArgumentOutOfRange_ConsoleColor))
    End Select
    
    If Width = 0 Or Height = 0 Then Exit Sub
    
    Dim FillByte As Byte
    Select Case VarType(FillChar)
        Case vbString
            If Len(FillChar) = 0 Then Exit Sub
            FillByte = Asc(FillChar)
        
        Case vbLong, vbInteger, vbByte:     FillByte = FillChar
        Case Else
            Throw Cor.NewArgumentException("Invalid fill Character", "FillChar")
    End Select

    If Len(FillChar) > 0 Then
        Dim TextColor As Long
        TextColor = (BackColor * &H10) Or Forecolor
        
        Dim ClearCoord As COORD
        ClearCoord.X = Left
        
        Dim i As Long
        For i = 0 To Height - 1
            ClearCoord.Y = Top + i
            If FillConsoleOutputCharacter(mOutputHandle, FillByte, Width, ByVal AsLong(ClearCoord), 0) = BOOL_FALSE Then IOError Err.LastDllError
            If FillConsoleOutputAttribute(mOutputHandle, TextColor, Width, ByVal AsLong(ClearCoord), 0) = BOOL_FALSE Then IOError Err.LastDllError
        Next i
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Call Init
End Sub

Private Sub Class_Terminate()
    Call SetConsoleCtrlHandler(AddressOf ControlBreakHandler, False)
    If mOwnsConsole Then Call FreeConsole
End Sub

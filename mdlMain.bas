Attribute VB_Name = "mdlMain"
Option Explicit
'> Characters information variables
    Public Type Character
        Data As String
        Name As String
        locX As String
        locY As String
        GameState As String
        Exits As String
    End Type
    Public Char(200) As Character
'< Characters information variables
'> Always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Global Const SWP_NOMOVE = 2
    Global Const SWP_NOSIZE = 1
    Global Const WndFlags = SWP_NOMOVE Or SWP_NOSIZE
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
'< Always on top
'> Constants
Public Const RET As String = vbCrLf
'< Constants
'> INI file
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'< INI file

Function GetIni(Class As String, VarToGet As String, IniFileName As String) As String
Dim RET As Long
Dim Temp As String * 2050
Dim lpAppName As String, lpKeyName As String, lpDefault As String, lpFileName As String

IniFileName = App.Path & "\" & IniFileName & ".ini"
lpAppName = Class 'Class name ( [User] )
lpDefault = IniFileName 'Ini location and file name
lpFileName = IniFileName 'Ini location and file name
RET = GetPrivateProfileString(lpAppName, VarToGet, IniFileName, Temp, Len(Temp), IniFileName)
GetIni = Mid(Temp, 1, RET)
If GetIni = IniFileName Then GetIni = "Error" Else _
  GetIni = Replace(GetIni, "<r>", vbCrLf)
End Function
Sub PutIni(Class As String, VarToPut As String, Value As String, IniFileName As String)
    Dim lpAppName As String, lpFileName As String, lpKeyName As String, lpString As String
    Dim RET As Long
    lpAppName = Class 'Class name ( [User] )
    lpKeyName = VarToPut 'Variable
    lpString = Value 'Value
    lpFileName = App.Path & "\" & IniFileName & ".ini" 'Ini location and file name
    RET = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
    If RET = 0 Then
      Beep
    End If
End Sub
'Function that checks if something is a part of an array
Function IsMember(ArrayName() As String, LookFor As String) As Boolean
'Entrance: ArrayName to look in and String to look for
    Dim I As Integer
    IsMember = False
    For I = 0 To UBound(ArrayName)
        If ArrayName(I) = LookFor Then IsMember = True: Exit For
    Next I
'Exit: True if found, false if not.
End Function
Public Function Proper(ByVal Str As String) As String
    '[Description]
    'Rudimentary routine to convert a mixed
    '     case string to Sentence case
    '[Declarations]
    Dim flgNextUpper As Boolean 'Is next alpha character the start of a new sentence
    Dim intIndex As Integer 'Current character being tested
    '[Code]
    Str = LCase(Str)
    flgNextUpper = True
    For intIndex = 1 To Len(Str)
        If Mid(Str, intIndex, 1) >= "a" _
          And Mid(Str, intIndex, 1) <= "z" _
          And flgNextUpper Then
        'Convert the current character
            Mid(Str, intIndex, 1) = UCase(Mid(Str, intIndex, 1))
        flgNextUpper = False
    End If
    If InStr(".!?:", Mid(Str, intIndex, 1)) Then
        'End of sentence reached
        flgNextUpper = True
    End If
Next 'character In String
Proper = Str
End Function
Function GetWord(Number As Integer, ByVal Word As String) As String
Dim I%, LastSpace%
    LastSpace = 1
    Word = " " & Word & " "
    For I = 1 To Number
        GetWord = Mid(Word, InStr(LastSpace, Word, " ") + 1, InStr(LastSpace + 1, Word, " ") - InStr(LastSpace, Word, " ") - 1)
        LastSpace = InStr(LastSpace + 1, Word, " ")
    Next I
End Function
'Increase sub
Sub Inc(Num As Integer)
    Num = Num + 1
End Sub
'Decrease sub
Sub Dec(Num As Integer)
    Num = Num - 1
End Sub

Function Ansi(Color As Integer) As String
    Select Case Color
    Case Is <= 7
        Ansi = Chr(27) & "[0m" & Chr(27) & "[3" & Color & "m"
    Case Is >= 8
        Ansi = Chr(27) & "[1m" & Chr(27) & "[3" & Color - 8 & "m"
    End Select
'ANSI Name  Code Sequence   Action
'
'*******************************************************************************
'
'     ED    ESC[ Pn J       Erases all or part of a display.
'                           Pn=0: erases from active position to end of display.
'                           Pn=1: erases from the beginning of display to
'                                 active position.
'                           Pn=2: erases entire display.
'
'     EL    ESC[ Pn K       Erases all or part of a line.
'                           Pn=0: erases from active position to end of line.
'                           Pn=1: erases from beginning of line to active
'                                 position.
'                           Pn=2: erases entire line.
'
'     ECH   ESC[ Pn X       Erases Pn characters.
'
'     CBT   ESC[ Pn Z       Moves active position back Pn tab stops.
'
'     SU    ESC[ Pn S       Scroll screen up Pn lines, introducing new blank
'                           lines at bottom.
'
'     SD    ESC[ Pn T       Scroll screen down Pn lines, introducing new blank
'                           lines at top.
'
'     CUP   ESC[ P1;P2 H    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     HVP   ESC[ P1;P2 f    Moves cursor to location P1 (vertical)
'                           and P2 (horizontal).
'
'     CUU   ESC[ Pn A       Moves cursor up Pn number of lines.
'
'     CUD   ESC[ Pn B       Moves cursor down Pn number of lines.
'
'     CUF   ESC[ Pn C       Moves cursor Pn spaces to the right.
'
'     CUB   ESC[ Pn D       Moves cursor Pn spaces backward.
'
'     HPA   ESC[ Pn '       Moves cursor to column given by Pn.
'
'     HPR   ESC[ Pn a       Moves cursor Pn characters to the right.
'
'     VPA   ESC[ Pn d       Moves cursor to line given by Pn.
'
'     VPR   ESC[ Pn e       Moves cursor down Pn number of lines.
'
'     IL    ESC[ Pn L       Inserts Pn new, blank lines.
'
'     ICH   ESC[ Pn @       Inserts Pn blank places for Pn characters.
'
'     DL    ESC[ Pn M       Deletes Pn lines.
'
'     DCH   ESC[ Pn P       Deletes Pn number of characters.
'
'     CPL   ESC[ Pn F       Moves cursor to beginning of line, Pn lines up.
'
'     CNL   ESC[ Pn E       Moves cursor to beginning of line, Pn lines down.
'
'     SGR   ESC[ Pn m       Changes display mode.
'                           Pn=0: Resets bold, blink, blank, underscore, and
'                           reverse.
'                           Pn=1: Sets bold (light_color).
'                           Pn=4: Sets underscore.
'                           Pn=5: Sets blink.
'                           Pn=7: Sets reverse video.
'                           Pn=8: Sets blank (no display).
'                           Pn=10: Select primary font.
'                           Pn=11: Select first alternate font.
'                           Pn=12: Select second alternate font.
'
'           ESC[ 2h         Lock keyboard. Ignores keyboard input until
'                           unlocked.
'
'           ESC[ 2i         Send screen to host.
'
'           ESC[ 2l         Unlock keyboard.
'
'           ESC[ 3 C m      Selects foreground colour C.
'
'           ESC[ 4 C m      Selects background colour C.
'
'                           C=0  Black
'                           C=1  Red
'                           C=2  Green
'                           C=3  Yellow
'                           C=4  Dark Blue
'                           C=5  Magenta
'                           C=6  Light Blue
'                           C=7  White
'
'****************************************************************************
'
'Here are the IBM / Ansi Characters

End Function
Function CountInString(ByVal sString As String, ByVal sToCount As String) As Long
    Dim lPos As Long
    
    CountInString = -1
    Do
        lPos = InStr(lPos + 1, sString, sToCount)
        CountInString = CountInString + 1
    Loop Until lPos = 0
End Function
Function ParseString(ByVal sString As String) As Variant
    Dim Numbers() As Long, lPos As Long, lPosEnd As Long, Counter As Long
    If CountInString(sString, "[") = 0 Then Exit Function
    ReDim Numbers(1 To CountInString(sString, "[")) As Long
    Do
        Counter = Counter + 1
        lPos = InStr(lPos + 1, sString, "[")
        If lPos = 0 Then Exit Do
        lPosEnd = InStr(lPos, sString, "]")
        Numbers(Counter) = Val(Mid(sString, lPos + 1, lPosEnd - lPos - 1))
    Loop
    ParseString = Numbers ' Variants can store arrays.
End Function


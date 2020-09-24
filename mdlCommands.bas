Attribute VB_Name = "mdlCommands"
Option Explicit
Sub DoCommands(Index As Integer, ByVal DataR$)
        Dim Command As String
        Dim Arguement As String
        Dim Transmit As String
        Dim I As Integer
        Dim Data As String
        Data = DataR
        DoEvents
'> Resolves the command from everything else
        Char(Index).Data = ""
        If InStr(Data, " ") Then
            Command = Left(Data, InStr(Data, " ") - 1)
            Arguement = Right(Data, Len(Data) - Len(Command) - 1)
        Else
            Command = Data
        End If
        If Left(Data, 1) = "'" Or Left(Data, 1) = ";" Then
            Command = Left(Data, 1)
            Arguement = Right(Data, Len(Data) - 1)
        End If
'> Resolves the command from everything else
'Checking aliases
        If IsMember(AliasList(), Command) Then Command = AliasToFull(Command)
'> Checking exists and moving if needed
        If InStr(Char(Index).Exits, "[" & Command & "]") Then
            Dim NFutureExit As String
            TransmitLocal Index, Char(Index).Name & " leaves " & Command & "."
            NFutureExit = FutureExit(Index, Command)
            RemovePC Index
            Char(Index).locY = GetExitY(NFutureExit)
            Char(Index).locX = GetExitX(NFutureExit)
            AddPC Index
            Char(Index).Exits = CurrentExits(Index)
            CommandLook Index
            TransmitLocal Index, Char(Index).Name & " arrives."
            Exit Sub
        End If
'< Checking exists and moving if needed
'> Other commands
        Select Case Command
' COMMAND: "Say"
        Case "say", "'"
            Send Index, "You say, '" & Arguement & "'"
            TransmitLocal Index, Char(Index).Name & " says, '" & Arguement & "'"
' COMMAND: "Emote"
        Case "emote", ";"
            Send Index, "-> " & Char(Index).Name & " " & Arguement
            TransmitLocal Index, "-> " & Char(Index).Name & " " & Arguement
' COMMAND: "Look"
        Case "look", "l"
            If Arguement = "" Then
                CommandLook (Index)
            Else 'Look AT something
                Dim NameIndex As Integer
                NameIndex = NameIsHere(Index, LCase(Arguement))
                If NameIndex > -1 Then
                    Send Index, StrConv(HeShe(NameIndex, "HeShe"), vbProperCase) & " is a " & _
                      GetIni(Char(NameIndex).Name, "Gender", "Users") _
                      & " " & GetIni(Char(NameIndex).Name, "Race", "Users") & "."
                    Send NameIndex, Char(Index).Name & " looks at you"
                    TransmitLocal Index, Char(Index).Name & " looks at " & Char(NameIndex).Name, NameIndex
                Else
                    Send Index, "It is not here"
                End If
            End If
'COMMANDS: Ready emotes
        Dim ToSelf$, ToOthers$, ToTarget$
        Case GetIni(Command, "ID", "Emotes"), GetIni(Command & "+", "ID", "Emotes")
            If NameIsHere(Index, LCase(GetWord(1, Arguement))) < 0 Or GetIni(Command & "+", "Self", "Emotes") = "Error" Then
                ToSelf = "You " & GetIni(Command, "Self", "Emotes")
                ToSelf = OpenTags(ToSelf, Index, Arguement)
                ToOthers = Char(Index).Name & " " & GetIni(Command, "Others", "Emotes")
                ToOthers = OpenTags(ToOthers, Index, Arguement)
                Send Index, ToSelf
                TransmitLocal Index, ToOthers
            Else
                ToSelf = "You " & GetIni(Command & "+", "Self", "Emotes")
                ToSelf = OpenTags(ToSelf, Index, Proper(Arguement))
                ToTarget = Char(Index).Name & " " & GetIni(Command & "+", "Target", "Emotes")
                ToTarget = OpenTags(ToTarget, Index, Replace(Arguement, GetWord(1, Arguement), "", , 1))
                ToOthers = Char(Index).Name & " " & GetIni(Command & "+", "Others", "Emotes")
                ToOthers = OpenTags(ToOthers, Index, Proper(Arguement))
                Send Index, ToSelf
                Send NameIsHere(Index, LCase(GetWord(1, Arguement))), ToTarget
                TransmitLocal Index, ToOthers, NameIsHere(Index, LCase(GetWord(1, Arguement)))
            End If
'COMMAND: "Help"
        Case "help"
            Send Index, "List of commands below: " & RET & RET & _
              "`say` or `'` - Use them to SAY things: say hi" & RET & _
              "`emote` or `;` - Use them to EMOTE things: emote smiles" & RET & _
              "`look` or `l` - Use them to LOOK around: look" & RET & _
              "`news` - Use it to see the last changed in the mud, changes without an asterik before them are in my todo list: news" & RET & _
              "`quit` - Use it to QUIT: quit"
'COMMAND: "Help"
        Case "news"
            On Error GoTo Error
            Dim News As String
            Open App.Path & "/" & "todo.dat" For Input As #1
            Do While Not EOF(1)
                Line Input #1, News
                If InStr(News, "=") Then News = Mid(News, InStr(News, "=") + 1, Len(News))
                News = Replace(News, "*-", "Done: ")
                News = Replace(News, "-", Ansi(15) & "Todo" & Ansi(7) & ": ")
                If News <> "" And (Left(News, 4) = "Done" Or Mid(News, 10, 4) = "Todo") Then Send Index, News
            Loop
Error:      Close #1
            If Err.Number > 0 Then Debug.Print Err.Description: Resume
' COMMAND: "Quit"
        Case "quit"
            Send Index, "See you later ;)"
            frmMain.wskAccept(Index).Close
            frmMain.CloseConnection Index
        Case Else
            Send Index, "What?"
            Debug.Print Command
        End Select
'< Other commands
End Sub
'The look command
Sub CommandLook(Index As Integer)
    Dim I As Integer
    Send Index, CurrentDesc(Index) & vbCrLf
    For I = frmMain.wskAccept.LBound To frmMain.wskAccept.UBound
        If Char(I).locX = Char(Index).locX And Char(I).locY = Char(Index).locY And I <> Index Then _
          Send Index, Char(I).Name & " is here."
    Next I
End Sub

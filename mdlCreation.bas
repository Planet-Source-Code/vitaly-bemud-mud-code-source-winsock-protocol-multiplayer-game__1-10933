Attribute VB_Name = "mdlCreation"
Option Explicit
'Character creation
Sub DoCreation(Index As Integer, DataR$)
    Dim LowData As String
    Dim Data As String
    Data = DataR
    Char(Index).Data = ""
    DoEvents
    LowData = LCase(Data)
    Select Case Char(Index).GameState
    Case "Name"
        If InStr(Data, " ") = 0 And Len(Data) > 0 Then
            Char(Index).Name = Proper(Data)
            frmMain.lstUsers.List(Index) = Index & ") " & Char(Index).Name
            If GetIni(Char(Index).Name, "OOCName", "Users") = "Error" Then
                Send Index, "This name doesn't exist. Are you sure '" & Char(Index).Name & "' would be a good name?"
                Char(Index).GameState = "NameConfirm"
            Else
                Send Index, "This character already exists. Enter your password, please."
                Char(Index).GameState = "PasswordCheck"
            End If
        Else
            Send Index, "Illegal name. Type again:"
        End If
    Case "NameConfirm"
        Select Case LowData
        Case "y", "yes"
            Send Index, "Please choose a password"
            Char(Index).GameState = "PasswordChoosing"
        Case "n", "no"
            Send Index, "Alright, choose a new one then"
            Char(Index).GameState = "Name"
        Case Else
            Send Index, "Yes or no, please"
        End Select
    Case "PasswordChoosing"
        PutIni Char(Index).Name, "OOCName", Char(Index).Name, "Users"
        PutIni Char(Index).Name, "Password", Data, "Users"
        Send Index, "Retype the password again"
        Char(Index).GameState = "PasswordConfirm"
    Case "PasswordConfirm"
        If Data = GetIni(Char(Index).Name, "Password", "Users") Then
            Send Index, "Password confirmed"
            Send Index, "Please choose a race: human/elf/gnome/kendar/dwarf"
            Char(Index).GameState = "Race"
        Else
            Send Index, "The passwords don't match. Please choose your password again."
            Char(Index).GameState = "PasswordChoosing"
        End If
    Case "Race"
        Select Case LowData
        Case "human", "gnome", "elf", "kendar", "dwarf"
            PutIni Char(Index).Name, "Race", LowData, "Users"
            Send Index, "You are now " & Ansi(15) & LowData
            Send Index, "I can't see very well here, are you male or female?"
            Char(Index).GameState = "Gender"
        Case Else
            Send Index, "This race must be from some other mud... Choose again: human/elf/gnome/kendar/dwarf"
        End Select
    Case "Gender"
        Select Case LowData
        Case "male", "m"
            PutIni Char(Index).Name, "Gender", "male", "Users"
            CreationFinish (Index)
        Case "female", "f"
            PutIni Char(Index).Name, "Gender", "female", "Users"
            CreationFinish (Index)
        Case Else
            Send Index, "Yet, pretend to be one of these: Male/Female"
        End Select
    Case "PasswordCheck"
        If Data = GetIni(Char(Index).Name, "Password", "Users") Then
            Char(Index).GameState = "Game"
            CreationFinish (Index)
        Else
            Send Index, "Wrong password. Enter your name again."
            Char(Index).GameState = "Name"
        End If
    End Select
End Sub
Sub CreationFinish(Index As Integer)
            If Char(Index).GameState = "Gender" Then Send Index, "Flesh grows in the right parts of your body as you choose a gender" & RET
            Char(Index).GameState = "Game": Char(Index).locX = 1: Char(Index).locY = 1: AddPC (Index)
            Send Index, "Welcome to BeMUD, " & Char(Index).Name & vbCrLf  'Creation stage is over
            Send Index, "Type " & Ansi(15) & "HELP" & Ansi(7) & " to get the list of the commands" & RET
            Send Index, CurrentDesc(Index)
            Char(Index).Exits = CurrentExits(Index)
            TransmitLocal Index, Char(Index).Name & " wakes up, yawning."
End Sub

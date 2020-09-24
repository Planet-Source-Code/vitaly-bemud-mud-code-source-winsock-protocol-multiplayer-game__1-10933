Attribute VB_Name = "mdlMudRelated"
Option Explicit
Public AliasList(3) As String
'Sending the data to the Index
Sub Send(Index As Integer, Data As String, Optional Color As Integer = 7)
Dim Wrap
    On Error Resume Next
    Wrap = 70
    If Len(Data) < Wrap Then
        frmMain.wskAccept(Index).SendData (Ansi(Color) & Data & vbCrLf)
    'Wrapping the text if it is too long.
    Else
        Dim Wrapped As String
        Wrapped = ""
        Do Until InStr(Wrap, Data, " ") = 0
            'Checks for RET in the string and wraps it by RETs and Spaces.
            If InStr(Data, vbCrLf) And InStr(Data, vbCrLf) < InStr(Wrap, Data, " ") Then
                Wrapped = Wrapped & Ansi(Color) & Left(Data, InStr(Data, vbCrLf))
                Data = Right(Data, Len(Data) - InStr(Data, vbCrLf))
            Else
                Wrapped = Wrapped & Ansi(Color) & Left(Data, InStr(Wrap, Data, " ") - 1) & vbCrLf
                Data = Right(Data, Len(Data) - InStr(Wrap, Data, " "))
            End If
        Loop
        Wrapped = Wrapped & Ansi(Color) & Data & vbCrLf
        frmMain.wskAccept(Index).SendData Wrapped
    End If
    DoEvents
End Sub
'Transmits the message to people in the same room with the the transmitor
Sub TransmitLocal(Index As Integer, Transmit As String, Optional NoSend As Integer = -1)
    Dim I As Integer
    Dim Arr As Variant
    Arr = ParseString(PCs(Index))
    If UBound(Arr) > 1 Then
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> Index And Arr(I) <> NoSend Then Send Val(Arr(I)), Transmit
        Next I
    End If
    End Sub
Function NameIsHere(Index As Integer, Name As String) As Integer
    Dim I As Integer
    Dim Arr As Variant
    NameIsHere = -1
    Arr = ParseString(PCs(Index))
    If UBound(Arr) > 1 Then
        For I = LBound(Arr) To UBound(Arr)
            If Arr(I) <> Index And LCase(Char(Arr(I)).Name) = Name Then NameIsHere = Arr(I)
        Next I
    End If
End Function
Function HeShe(Index%, GenderType$) As String
    If GetIni(Char(Index).Name, "Gender", "Users") = "male" Then
        Select Case GenderType
        Case "HeShe"
            HeShe = "he"
        Case "HisHer"
            HeShe = "his"
        End Select
    Else
        Select Case GenderType
        Case "HeShe"
            HeShe = "she"
        Case "HisHer"
            HeShe = "her"
        End Select
    End If
End Function
Function OpenTags(ByVal Str$, Index%, Arguement$) As String
    If NameIsHere(Index, LCase(GetWord(1, Arguement))) > -1 Then Str = Replace(Str, "<target>", Char(NameIsHere(Index, LCase(GetWord(1, Arguement)))).Name): _
      Arguement = Trim(Replace(Arguement, GetWord(1, Arguement), "", , 1))
    Str = Replace(Str, "<arg>", Arguement)
    Str = Replace(Str, "<hisher>", HeShe(Index, "HisHer"))
    OpenTags = Str
End Function
Sub AliasInput()
'> Set GlobalAliases
    AliasList(0) = "n"
    AliasList(1) = "s"
    AliasList(2) = "w"
    AliasList(3) = "e"
'< Set GlobalAliases
End Sub
Function AliasToFull(Alias As String) As String
    Select Case Alias
    Case "n"
        AliasToFull = "north"
    Case "s"
        AliasToFull = "south"
    Case "w"
        AliasToFull = "west"
    Case "e"
        AliasToFull = "east"
    End Select
End Function

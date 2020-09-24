VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "WINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "BeMUD"
   ClientHeight    =   3000
   ClientLeft      =   8340
   ClientTop       =   1455
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5400
   Begin VB.TextBox txtOutput 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.ListBox lstUsers 
      Height          =   2985
      ItemData        =   "frmMain.frx":0000
      Left            =   3960
      List            =   "frmMain.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2700
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock wskAccept 
      Index           =   0
      Left            =   4200
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskListen 
      Left            =   3720
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users Management"
      Visible         =   0   'False
      Begin VB.Menu mnuUserIP 
         Caption         =   "IP"
      End
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastConnection As Integer
Private Sub Form_DblClick()
    Clipboard.SetText "bemud.selfhost.com" & ":" & wskListen.LocalPort
End Sub
Private Sub Form_Load()
    Dim I As Integer
'> Sets the window topmost
    I = SetWindowPos(Me.hWnd, HWND_TOPMOST, _
    Me.Left \ Screen.TwipsPerPixelX, Me.Top \ Screen.TwipsPerPixelY, _
    Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 0)
'< Sets the window topmost
    frmMain.Caption = "BeMUD - BeMud.selfhost.com"
    AliasInput
'> Start to listen
    wskListen.LocalPort = 23
    wskListen.Listen
'< Start to listen
End Sub
'This sub keeps the objects on the form in the right size
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        lstUsers.Left = ScaleWidth - lstUsers.Width
        lstUsers.Height = ScaleHeight
        txtOutput.Width = ScaleWidth - lstUsers.Width - 50
        txtOutput.Height = ScaleHeight - txtInput.Height - 50
        txtInput.Width = ScaleWidth - lstUsers.Width - 50
        txtInput.Top = ScaleHeight - txtInput.Height
    End If
End Sub

Private Sub Form_Terminate()
Dim I%, J%, Imin%, Imax%, Jmin%, Jmax%
    Imin = GetIni("Area", "MinX", "Void")
    Jmin = GetIni("Area", "MinY", "Void")
    Imax = GetIni("Area", "MaxX", "Void")
    Jmax = GetIni("Area", "MaxY", "Void")
    For I = Imin To Imax
        For J = Jmin To Jmax
            PutIni "Room " & I & "," & J, "PCs", "", "Void"
        Next J
    Next I
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuUsers
End Sub
Private Sub mnuKickUser_Click()
Dim I As Integer
    For I = lstUsers.ListCount - 1 To 0 Step -1
        If lstUsers.Selected(I) = True Then
            lstUsers.Selected(I) = False
            Call Log("User kicked", Char(lstUsers.ItemData(I)).Name)
            wskAccept(lstUsers.ItemData(I)).Close
            CloseConnection (lstUsers.ItemData(I))
        End If
    Next I
End Sub

Private Sub mnuUserIP_Click()
Dim I As Integer
    For I = 0 To lstUsers.ListCount - 1
        If lstUsers.Selected(I) = True Then Call Log(wskAccept(lstUsers.ItemData(I)).RemoteHostIP, Char(lstUsers.ItemData(I)).Name)
    Next I
End Sub

'Sends the text from the server side to selected clients
Private Sub txtInput_KeyPress(KeyAscii As Integer)
Dim I As Integer
    If KeyAscii = 13 Then
        For I = 0 To lstUsers.ListCount - 1
            If lstUsers.Selected(I) = True Then Send lstUsers.ItemData(I), txtInput.Text, 12
        Next I
        txtInput.SelStart = 0
        txtInput.SelLength = Len(txtInput.Text)
        KeyAscii = 0
    End If
End Sub
'Keeps the output scrollbar down-most
Private Sub txtOutput_Change()
    txtOutput.SelStart = Len(txtOutput.Text)
End Sub
Private Sub wskAccept_Close(Index As Integer)
    CloseConnection Index
End Sub
Private Sub wskAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Quo(256) As Integer
    Dim QuoText(256, 10) As String
    Dim PartData As String
    wskAccept(Index).GetData PartData
    'Waits till there is an ENTER in the recieved data
Do
    Debug.Print "Data arrived: " & Timer
    DoEvents
    If Quo(Index) > 0 Then PartData = QuoText(Index, Quo(Index)): Dec Quo(Index)
    If Asc(PartData) <> 8 Then _
      Char(Index).Data = Char(Index).Data & PartData _
      Else If Len(Char(Index).Data) > 0 Then _
        Char(Index).Data = Left(Char(Index).Data, Len(Char(Index).Data) - 1)
    If InStr(Char(Index).Data, vbCrLf) Or InStr(Char(Index).Data, vbLf) Then
'> Removes the enter
        If InStr(Char(Index).Data, vbCrLf) Then
            If Right(Char(Index).Data, Len(Char(Index).Data) - InStr(Char(Index).Data, vbCrLf) - 1) <> "" Then
                Inc Quo(Index)
                QuoText(Index, Quo(Index)) = Right(Char(Index).Data, Len(Char(Index).Data) - InStr(Char(Index).Data, vbCrLf) - 1)
            End If
            Char(Index).Data = Left(Char(Index).Data, InStr(Char(Index).Data, vbCrLf) - 1)
        End If
        If InStr(Char(Index).Data, vbLf) Then
            If Right(Char(Index).Data, Len(Char(Index).Data) - InStr(Char(Index).Data, vbLf)) <> "" Then
                Inc Quo(Index)
                QuoText(Index, Quo(Index)) = Right(Char(Index).Data, Len(Char(Index).Data) - InStr(Char(Index).Data, vbLf))
            End If
            Char(Index).Data = Left(Char(Index).Data, InStr(Char(Index).Data, vbLf) - 1)
        End If
'< Removes the enter
        Log Char(Index).Data, Char(Index).Name
'> Checking to what Stage the command was sent
'Stages: 3 - Game mode, 1 - Creation mode
        Select Case Char(Index).GameState
        Case "Game"
            DoCommands Index, Char(Index).Data
        Case "Gender", "Name", "NameConfirm", "PasswordCheck", "PasswordChoosing", "PasswordConfirm", "Race"
            DoCreation Index, Char(Index).Data
        End Select
        Debug.Print "Data sent: " & Timer
    End If
'< Checking to what Stage the command was sent
Loop Until Quo(Index) = 0
End Sub
Private Sub wskAccept_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug.Print Error
    If Error = 10054 Then
    CloseConnection Index
    Debug.Print "Error " & Number & " - " & Description
    End If
End Sub
'Letting users to connect the mud
Private Sub wskListen_ConnectionRequest(ByVal requestID As Long)
    Call GetFreeWinsockIndex 'Looking for free Winsock
    wskAccept(LastConnection).Close
    Log "Accepting access on " & LastConnection & " Winsock.", "Server"
    wskAccept(LastConnection).Accept requestID
    lstUsers.AddItem LastConnection & ") Unknown"
    lstUsers.ItemData(lstUsers.NewIndex) = LastConnection
    Char(LastConnection).GameState = "Name"
    Char(LastConnection).Name = "Unknown"
    Send LastConnection, Chr(27) & "[2J"
    Send LastConnection, GetIni("Logo", "Draw", "Graphics"), 9
    Send LastConnection, RET & "Welcome to BeMUD, please enter your name: "
End Sub
'Searches for free Winsock
Sub GetFreeWinsockIndex()
    For LastConnection = wskAccept.LBound To wskAccept.UBound
        If wskAccept(LastConnection).State = sckClosed Then Exit Sub
    Next
    Load wskAccept(LastConnection)
End Sub
'Logging the data on the output server screen
Sub Log(Data As String, Sign As String)
    txtOutput.Text = txtOutput.Text & Sign & ": " & Data & vbCrLf
End Sub
'Closing connection and resetting the personal Vars
Sub CloseConnection(Index As Integer)
 Dim I As Integer
    If Char(I).GameState = "Game" Then TransmitLocal Index, Char(Index).Name & " yawns and goes to sleep."
    Log "Closing access on " & Index & " Winsock.", "Server"
    If Index = wskAccept.UBound And Index > 0 Then Unload wskAccept(Index) Else wskAccept(Index).Close
    Char(Index).Exits = ""
    Char(Index).Data = ""
    Char(Index).GameState = 0
    RemovePC Index
    Char(Index).locX = 0: Char(Index).locY = 0
    Char(Index).Name = ""
    For I = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(I) = Index Then lstUsers.RemoveItem I: Exit Sub
    Next I
End Sub

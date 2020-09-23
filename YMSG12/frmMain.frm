VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YMSG12 Client Example By: Matthew Robertson"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTest 
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton cmdViewShareFiles 
      Caption         =   "ViewShareFile()"
      Height          =   285
      Left            =   7920
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendIMV 
      Caption         =   "SendIMV()"
      Height          =   285
      Left            =   7920
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSendFile 
      Caption         =   "SendFile()"
      Height          =   285
      Left            =   6480
      TabIndex        =   19
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGotoUser 
      Caption         =   "GoToUser()"
      Height          =   285
      Left            =   6480
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox lstChatters 
      Height          =   2205
      Left            =   6360
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Height          =   1215
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton cmdChatSend 
      Caption         =   "Send Chat"
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtChatMsg 
      Height          =   495
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ComboBox cboAway 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton cmdSendPM 
      Caption         =   "Send"
      Height          =   405
      Left            =   5640
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtMsg 
      Height          =   405
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtPMWho 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox lstBuddies 
      Height          =   2205
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ComboBox cboProfiles 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock wskYMSG 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Robertson"
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   1640
      Width           =   2775
   End
   Begin VB.Line Line6 
      X1              =   6360
      X2              =   9240
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Shit:"
      Height          =   195
      Index           =   7
      Left            =   6600
      TabIndex        =   29
      Top             =   120
      Width           =   750
   End
   Begin VB.Line Line5 
      X1              =   6360
      X2              =   6360
      Y1              =   1960
      Y2              =   120
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room List"
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   28
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Chat:"
      Height          =   195
      Index           =   5
      Left            =   3360
      TabIndex        =   27
      Top             =   2040
      Width           =   885
   End
   Begin VB.Line Line4 
      X1              =   9240
      X2              =   3120
      Y1              =   1960
      Y2              =   1960
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Simple PM:"
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   26
      Top             =   840
      Width           =   795
   End
   Begin VB.Line Line3 
      X1              =   3120
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   120
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   3120
      Y1              =   120
      Y2              =   4560
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buddies:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profiles:"
      Height          =   195
      Index           =   2
      Left            =   3360
      TabIndex        =   24
      Top             =   0
      Width           =   555
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo ID:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   0
      Width           =   720
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Robertson"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1220
      Width           =   1675
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Buddie_Off(OffBud As String)
On Error Resume Next
' loops threw buddies list 'till it finds the online buddie who wents off
With lstBuddies
 For i = 0 To .ListCount - 1
  If LCase(.List(i)) = LCase(OffBud & " (online)") Then
    .List(i) = OffBud
    Exit Sub
  End If
 Next
End With
End Sub


Sub Buddie_On(OnBud As String)
On Error Resume Next
'loops threw buddies 'till it finds the buddie to display as onlne
With lstBuddies
 For i = 0 To .ListCount - 1
  If LCase(.List(i)) = LCase(OnBud) Then
    .List(i) = OnBud & " (Online)"
    Exit Sub
  End If
 Next
.AddItem OnBud & " (Online)" ' adds them if they r not on the list for w/e reason, like if u just added em"
End With
End Sub


Sub Chatter_Join(Chatter As String)
Chatter_Part Chatter ' to prevent the name from being on the list twise
lstChatters.AddItem Chatter
End Sub

Sub Chatter_Part(Chatter As String)
On Error Resume Next
With lstChatters
 For i = 0 To .ListCount - 1
    If LCase(.List(i)) = LCase(Chatter) Then .RemoveItem i
 Next
End With
End Sub


Sub Incomming(ByVal Data As String)
On Error Resume Next
' better then having all this code in winsock
' only handels a few things, there are a LOT more if u want to make afully functional chat client
Dim sptData() As String
sptData = Split(Data, "À€")
Select Case Asc(Mid(Data, 12, 1))
 Case 168 ' chat text
    AddText sptData(3) & ": " & sptData(5)
 Case 155 ' user departs room
    If sptData(5) = "109" Then sptData(5) = sptData(6)
    Chatter_Part sptData(5)
 Case 152 ' user joins room
    Incomming_ChatList Data
 Case 150 ' join room
    lstChatters.Clear
    Incomming_RoomJoin sptData(1)
 Case 87 ' logging in
    SendPack PostLogin(YMSG.ID, YMSG.PW, sptData(3))
    YMSG.Key = Mid(Data, 17, 4) ' may b used later
 Case 85 ' logged in
    GetProfiles YMSG.ID, sptData(5), cboProfiles
    Incomming_BuddieList sptData(1)
    cmdLogin.Caption = "Logout"
    lblStat = "Logged in"
    cboAway.Text = "-Available"
 Case 6 ' incomming pm
    AddText "PM to: " & sptData(1) & " from: " & sptData(3) & " : " & sptData(7)
 Case 2 ' offline buddie
    Buddie_Off sptData(1)
 Case 1 ' online buddie
    Incomming_BuddieListOnline Data
End Select
End Sub

Sub AddText(Txt As String)
' i might change this to richtext sometime
With txtIn
    .SelStart = Len(.Text)
    If Not .Text = "" Then .SelText = vbCrLf ' line break
    .SelText = FilterYahooText(Txt)
    .SelStart = Len(.Text) ' scroll down
End With
End Sub

Function FilterYahooText(ByVal Str As String)
On Error GoTo Error
' filter out everyhtign between < and >, and  and m
Dim i As Integer, ii As Integer, Llp As Boolean
Llp = True
For lp = 1 To 12 ' better then a loop because this wont ever loop forever
 Llp = False
  i = InStr(Str, "<")
 If Not i = 0 Then
    ii = InStr(i, Str, ">")
    If Not ii = 0 Then
     Str = Left(Str, i - 1) & Right(Str, Len(Str) - ii)
     Llp = True
    End If
 End If
  i = InStr(Str, "[")
 If Not i = 0 Then
    ii = InStr(i, Str, "m")
    If Not ii = 0 Then
     Str = Left(Str, i - 1) & Right(Str, Len(Str) - ii)
     Llp = True
    End If
 End If
    DoEvents
If Llp = False Then Exit For
Next
Error:
FilterYahooText = Str
End Function
Sub Incomming_BuddieList(Buddies As String)
On Error Resume Next
' gets the overall buddies list
Dim Bud() As String
Bud = Split(Replace(Buddies, Chr(&HA), ","), ",")
For i = 0 To UBound(Bud)
    If InStr(Bud(i), ":") Then Bud(i) = Mid(Bud(i), InStr(Bud(i), ":") + 1)
    If Not Bud(i) = "" Then lstBuddies.AddItem Trim(Bud(i))
Next
End Sub


Sub Incomming_BuddieListOnline(Data As String)
On Error Resume Next
' get online buddie(s)
Dim Bud() As String, n As Integer
Bud = Split(Data, "À€7À€")
For i = 1 To UBound(Bud) ' incase its more then 1 person at a time
    n = InStr(Bud(i), "À€")
    If n > 1 Then Bud(i) = Left(Bud(i), n - 1)
    Buddie_On Bud(i)
Next
End Sub

Sub Incomming_ChatList(Data As String)
On Error Resume Next
Dim Chatter() As String, n As Integer
Chatter = Split(Data, "À€109À€")
For i = 1 To UBound(Chatter)
    n = InStr(Chatter(i), "À€")
    If n > 1 Then Chatter(i) = Left(Chatter(i), n - 1)
    Chatter_Join Chatter(i)
Next
End Sub

Sub Incomming_RoomJoin(ID As String)
SendPack JoinChat(ID, YMSG.Room, YMSG.Key)
End Sub

Sub LoadInfo()
txtID = GetSetting("YMSG12", "Login", "ID", "")
txtPW = GetSetting("YMSG12", "Login", "PW", "")
txtRoom = GetSetting("YMSG12", "Chat", "Room", "")
End Sub

Sub SaveInfo()
If Not YMSG.ID = "" Then SaveSetting "YMSG12", "Login", "ID", YMSG.ID
If Not YMSG.PW = "" Then SaveSetting "YMSG12", "Login", "PW", YMSG.PW ' u might wonna add mroe surcurty
If Not YMSG.Room = "" Then SaveSetting "YMSG12", "Chat", "Room", YMSG.Room
End Sub

Function SendPack(Packet As String) As Boolean
' type just sendpack blag
' or: if sendpack(blah) = false then error
On Error GoTo Error
 wskYMSG.SendData Packet
 Debug.Print "   " & Packet
 SendPack = True
 Exit Function
Error:
 SendPack = False
End Function

Sub Sleep(ByVal Sec As Long)
Sec = Timer & Sec
Do Until Timer > Sec
    DoEvents
Loop
End Sub

Private Sub cboAway_Click()
If cboAway.Text = "-Custom Msg" Then
    SendPack AwayMessage(InputBox("Custom away message:", "Away Message"))
ElseIf cboAway.Text = "-Available" Then
    SendPack AwayMessage("")
Else
    SendPack AwayMessage(cboAway.Text, 1)
End If
End Sub


Private Sub cmdAdd_Click()
SendPack AddBuddie(cboProfiles.Text, txtPMWho)
End Sub



Private Sub cmdChatSend_Click()
SendPack SendChat(cboProfiles.Text, txtChatMsg, YMSG.Room, YMSG.Key)
AddText cboProfiles.Text & ": " & txtChatMsg
txtChatMsg = ""
End Sub

Private Sub cmdGotoUser_Click()
SendPack GoToUser(cboProfiles.Text, txtTest)
End Sub

Private Sub cmdJoin_Click()
YMSG.Room = txtRoom
SendPack Prejoin(cboProfiles.Text)
End Sub

Private Sub cmdLogin_Click()
If cmdLogin.Caption = "Login" Then
    With YMSG
     .ID = txtID
     .PW = txtPW
     .Server = "scs.msg.yahoo.com"
     wskYMSG.Close
     wskYMSG.Connect .Server, 5050
    End With
    lblStat = "Connecting..."
ElseIf cmdLogin.Caption = "Logout" Then
    Call wskYMSG_Close
    cmdLogin.Caption = "Login"
    lblStat = "Logged out"
End If
End Sub



Private Sub cmdSendFile_Click()
SendPack SendFile(cboProfiles.Text, txtTest, "http://www.google.com/images/logo.gif", , "The World-Wide Web search engine that indexes the greatest number of web pages.")
End Sub

Private Sub cmdSendIMV_Click()
SendPack SendIMV(cboProfiles.Text, txtTest, "irobot;1")
End Sub


Private Sub cmdSendPM_Click()
SendPack SendPM(cboProfiles.Text, txtPMWho, txtMsg)
txtMsg = ""
End Sub





Private Sub cmdViewShareFiles_Click()
' from what i understand this dont work on yahoo messy 6
SendPack ViewShareFiles(cboProfiles.Text, txtTest)
End Sub

Private Sub Form_Load()
LoadInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveInfo
For i = 0 To Forms.Count - 1
    Unload Forms(i)
Next
End
End Sub


Private Sub lstBuddies_Click()
Dim Bud As String
Bud = lstBuddies.Text
If Right(Bud, 9) = " (Online)" Then Bud = Left(Bud, Len(Bud) - 9)
txtPMWho = Bud
End Sub

Private Sub lstChatters_Click()
txtPMWho = lstChatters.Text
End Sub


Private Sub txtChatMsg_Change()
cmdChatSend.Default = True
End Sub

Private Sub txtMsg_Change()
cmdSendPM.Default = True
End Sub

Private Sub wskYMSG_Close()
wskYMSG.Close
lstBuddies.Clear
cboProfiles.Clear
lstChatters.Clear
cmdLogin.Caption = "Login"
lblStat = "Disconencted!"
End Sub

Private Sub wskYMSG_Connect()
SendPack PreLogin(YMSG.ID)
End Sub

Private Sub wskYMSG_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
wskYMSG.GetData Data
Debug.Print Asc(Mid(Data, 12, 1)) & "- " & Data
Incomming Data
End Sub



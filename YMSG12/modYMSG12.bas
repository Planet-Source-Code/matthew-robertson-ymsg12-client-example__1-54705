Attribute VB_Name = "modYMSG12"
' modYMSG12 By: Matthew Robertson
' YMSG12ENCRYPT.dll by: ScriptedMind
'
' sniffed most packets off yahoo messy and yahelite
' feal free to use, but dont forget cridits!
'
' mailto:uphome@nbnet.nb.ca

Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG12ENCRYPT.dll" (ByVal username As String, ByVal password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Type typYMSG
    Server      As String
    ID          As String
    PW          As String
    Profiles(6) As String
    Room        As String
    RoomID      As String
    Key         As String
End Type
Global YMSG As typYMSG

Function AwayMessage(Msg As String, Optional Busy As Integer = 0)
If Msg = "Idle" Then Busy = 2
If LCase(Msg) = "invisible" Then
    AwayMessage = Packet("c5", "13À€2À€")
Else
    AwayMessage = Packet("c6", "10À€99À€19À€" & Msg & "À€47À€" & Busy & "À€187À€0À€")
End If
End Function

Function Prejoin(ID As String)
Prejoin = Packet(96, "109À€" & ID & "À€1À€" & ID & "À€6À€abcdeÀ€", YMSG.Key)
End Function
Function JoinChat(ID As String, Room As String, Key As String)
JoinChat = Packet(98, "1À€" & ID & "À€62À€À€2À€À€104À€" & Room & "À€", Key)
End Function
Function GoToUser(ID As String, Who As String)
GoToUser = Packet(97, "1À€" & ID & "À€109À€" & Who & "À€62À€2À€")
End Function
Function SendChat(From As String, Msg As String, Optional Room As String, Optional Key As String)
SendChat = Packet("A8", "1À€" & From & "À€104À€" & Room & "À€117À€" & Msg & "À€124À€1À€", Key)
End Function
Function SendPM(From As String, Who As String, Msg As String)
SendPM = Packet(17, "1À€" & From & "À€5À€" & Who & "À€14À€" & Msg & "À€97À€1À€")
End Function
Function SendFile(From As String, Who As String, URL As String, Optional Size As String = "Undefined", Optional Msg As String = "")
'sends a url as if it where a file transfer (the size can b a string)
Dim FileName As String
FileName = Right(URL, Len(URL) - InStrRev(URL, "/"))
SendFile = Packet("4D", "5À€" & Who & "À€49À€FILEXFERÀ€1À€" & From & "À€14À€" & Msg & "À€13À€1À€27À€" & FileName & "À€28À€" & Size & "À€20À€" & URL & "À€")
End Function
Function SendIMV(From As String, WhoTo As String, IMV As String)
SendIMV = Packet("4D", "49À€IMVIRONMENTÀ€1À€" & From & "À€14À€À€13À€0À€5À€" & WhoTo & "À€63À€" & IMV & "À€64À€0À€")
End Function
Function ViewShareFiles(From As String, Who As String)
ViewShareFiles = Packet("4D", "5À€" & Who & "À€49À€FILEXFERÀ€1À€" & From & "À€13À€5À€54À€MSG1.0À€")
End Function
Function AddBuddie(ID As String, Who As String, Optional Grp As String = "Buddies", Optional Msg As String)
AddBuddie = Packet(83, "1À€" & ID & "À€7À€" & Who & "À€14À€" & Msg & "À€65À€" & Grp & "À€")
End Function
Sub GetProfiles(MainID As String, Profiles As String, Optional Cbo As ComboBox)
'ymsg.profiles(num) will return that profile, but if there is no profiles it will return the main name
'not the best coding ever but it was the fastest way i could think to do it
On Error Resume Next
Dim Spt() As String, i As Integer
Spt = Split(Profiles & ",", ",")
i = UBound(Spt)
If i > 6 Then i = 6
With YMSG
 For i = 0 To i
    If Spt(i) = "" Or Left(Spt(i), 2) = "--" Then Exit For ' when somein fucks up
    .Profiles(i) = Spt(i)
    If Not Cbo Is Nothing Then Cbo.AddItem Spt(i) ' adds to a combo box if present
 Next
 For i = UBound(Spt) To 6 ' if u have all profiles this will do nothing
    .Profiles(i) = MainID
 Next
End With
If Not Cbo Is Nothing Then Cbo.Text = MainID
End Sub
Function PostLogin(ID As String, PW As String, SD As String)
Dim Enc(1) As String
On Error GoTo Error
Enc(0) = String(80, 0)
Enc(1) = String(80, 0)
'i think scriptedmind stoll the soruce to the DLL off gaim(gaim.sf.net) and deeps(yahelite.org) old DLL vbmod but yea... (im not 100% sure)
If YMSG12_ScriptedMind_Encrypt(ID, PW, SD, Enc(0), Enc(1), 1) = False Then
    'incase of error
    MsgBox "Error on: YMSG12ENCRYPT.DLL", vbCritical, "YMSG12ENCRYPT.DLL"
    GoTo Error
End If
 For i = 0 To 1
    Enc(i) = Left$(Enc(i), InStr(1, Enc(i), Chr(0)) - 1)
 Next
PostLogin = Packet(54, "6À€" & Enc(0) & "À€96À€" & Enc(1) & "À€0À€" & ID & "À€2À€" & ID & "À€192À€-1À€1À€" & ID & "À€135À€6,0,0,0000À€148À€360À€")
Exit Function
Error: ' error (53 means dll not found)
PostLogin = Err
End Function
Function PreLogin(ID As String)
PreLogin = Packet(57, "1À€" & ID & "À€")
End Function


Function Packet(PackType As String, Pack As String, Optional ByVal Key As String)
'adds header to packet
' i seen a lot of other codes where this was coded usng a 'calc size' function
' wich looped till the packlen was under 256 and counted the times it had to loop
' wich was simple dividing, and then the remaindure, wich can b done simply w/ 'mod'

If Key = "" Then Key = String(4, 0) ' key is just nothing in most cases
Packet = "YMSG" & Chr(0) & Chr(12) & String(2, 0) & _
Chr(Fix(Len(Pack) / 256)) & Chr(Len(Pack) Mod 256) & _
Chr(0) & Chr("&h" & PackType) & String(4, 0) & Key & _
Pack '  cleaner then most header functions :)
End Function


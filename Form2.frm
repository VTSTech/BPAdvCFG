VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "#BPAdvCFG - Burnout Paradise Advanced Config Chat! (irc.webchat.org)"
   ClientHeight    =   11715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13665
   LinkTopic       =   "Form2"
   ScaleHeight     =   11715
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   11280
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   10140
      Left            =   12000
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   315
      Left            =   12720
      TabIndex        =   4
      Top             =   10800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nick"
      Height          =   315
      Left            =   12720
      TabIndex        =   3
      Top             =   10440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   12000
      TabIndex        =   2
      Top             =   10440
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Text            =   "Enter text here. Press Send to .. send. Press Nick to change your nick name. ENTER will send :)"
      Top             =   10440
      Width           =   11775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   10215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
   Begin MSWinsockLib.Winsock IRC 
      Left            =   120
      Top             =   11280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Your Nickname:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3720
      TabIndex        =   9
      Top             =   11400
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4920
      TabIndex        =   8
      Top             =   11400
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   7920
      TabIndex        =   7
      Top             =   11400
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Connection State:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   6600
      TabIndex        =   6
      Top             =   11400
      Width           =   1275
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BPChan
Dim BPData
Dim BPNick
Dim x
Dim WhoIs
Dim Nick
Dim Names()

Private Sub Command1_Click()
Text1.Text = Text1.Text & vbCrLf & "<Chat:#BPAdvCFG>(" & BPNick & ":) " & Text2.Text & vbCrLf
IRC.SendData "PRIVMSG " & BPChan & " :" & Text2.Text & vbCrLf
End Sub

Private Sub Command2_Click()
Tmp = InputBox("What is your desired nickname ?")
IRC.SendData "NICK " & Tmp & vbCrLf
Label3.Caption = Tmp
End Sub

Private Sub Command3_Click()
Unload Form2
End Sub

Private Sub Form_Load()
Form2.Visible = True
Timer1.Interval = 500
Timer1.Enabled = True
IRC.Protocol = sckTCPProtocol
IRC.RemoteHost = "irc.webchat.org"
IRC.RemotePort = 6667
IRC.Connect
BPChan = "#BPAdvCFG"
ReDim Names(1)
End Sub

Private Sub IRC_Connect()
Randomize Timer
BPNick = "BPGamer" & Int(Rnd * 899) + 100
IRC.SendData "NICK " & BPNick & vbCrLf
Label3.Caption = BPNick
IRC.SendData "USER BPAdvCFG BPAdvCFG BPAdvCFG :Burnout Paradise Player" & vbCrLf
End Sub

Private Sub IRC_DataArrival(ByVal bytesTotal As Long)
IRC.GetData BPData, vbString
'On Error Resume Next

For x = 1 To Len(BPData)
'***Server Ping/Pong
If Mid$(BPData, x, 6) = "PING :" And Len(Pong) < 1 Then
Pong = Mid$(BPData, x + 6, Len(BPData) - 2)
IRC.SendData "PONG " & Pong & vbCrLf
IRC.SendData "JOIN " & BPChan & " " & ChanKey & vbCrLf
End If

If Len(Text1.Text) > 10000 Then
Text1.Text = ""
End If
':Veritas!NOPE@=b4i47-132-223-077.home1.cgocable.net PRIVMSG #bpadvcfg :testing
If Mid$(BPData, x, 20) = " PRIVMSG #bpadvcfg :" Then
Text1.Text = Text1.Text & "<Chat:#BPAdvCFG>(" & Nick & ":) " & Mid$(BPData, x + 20, Len(BPData) - x - 18)
':katana.webmaster.com 332 BPGamer100 #bpadvcfg :#BPAdvCFG - Burnout Paradise Advanced Config Tool v0.2.5 Download: http://www.mediafire.com/?2i9fdda6k0ctkot Homepage:  http://nigelt.wordpress.com/2012/03/04/bpadvcfg-burnout-paradise-advanced-config-tool/
ElseIf Mid$(BPData, x, 27) = " 332 " & BPNick & " #bpadvcfg :" Then
Text1.Text = Text1.Text & "<Topic:#BPAdvCFG>" & Mid$(BPData, x + 26, Len(BPData) - x - 23) & vbCrLf
ElseIf Mid$(BPData, x, 27) = "Nickname is already in use." Then
Tmp = InputBox("That nickname was taken! Please choose another.")
IRC.SendData "NICK " & Tmp & vbCrLf
Label3.Caption = Tmp
End If

':katana.webmaster.com 332 BPGamer414 #bpadvcfg :#BPAdvCFG - Burnout Paradise Advanced Config Tool v0.2.5 Download: http://www.mediafire.com/?2i9fdda6k0ctkot Homepage:  http://nigelt.wordpress.com/2012/03/04/bpadvcfg-burnout-paradise-advanced-config-tool/

':katana.webmaster.com 265 BPGamer17 :Current local users: 1880  Max: 3699
':katana.webmaster.com 266 BPGamer17 :Current global users: 13224  Max: 18893
':katana.webmaster.com 377 BPGamer17 z-default 1884 1328740714 :Last MOTD change information: Wed, 08 Feb 2012 14:38:34 -0800
':katana.webmaster.com 375 BPGamer17 :- katana.webmaster.com Message of the day
':katana.webmaster.com 372 BPGamer17 :- For the message of the day please visit: http://community.webmaster.com/motd/
':katana.webmaster.com 376 BPGamer17 :- End of message of the day
':katana.webmaster.com 221 BPGamer17 :+ixpemJMn
':BPGamer17!BPAdvCFG@=b4i47-132-223-077.home1.cgocable.net JOIN :#bpadvcfg
':katana.webmaster.com 332 BPGamer17 #bpadvcfg :#BPAdvCFG - Burnout Paradise Advanced Config Tool v0.2.5 Download: http://www.mediafire.com/?2i9fdda6k0ctkot Homepage:  http://nigelt.wordpress.com/2012/03/04/bpadvcfg-burnout-paradise-advanced-config-tool/
':katana.webmaster.com 333 BPGamer17 #bpadvcfg Veritas 1332572579
':katana.webmaster.com 353 BPGamer17 = #bpadvcfg :BPGamer17 @Veritas
':katana.webmaster.com 366 BPGamer17 #bpadvcfg :End of /NAMES list.

':arena.webmaster.com
'***Nick/Address Parsing
WhoIs = Split(BPData, "!")
Nick = Mid$(WhoIs(0), 2, Len(WhoIs(0)))
Next x

For y = 1 To UBound(Names)
If Names(y) = Nick Then
a = a
ElseIf Mid$(Nick, 1, 5) = "ING :" Then
a = a
ElseIf InStr(1, Nick, "webmaster") = 1 Or InStr(2, Nick, "webmaster") = 2 Or InStr(3, Nick, "webmaster") = 3 Or InStr(4, Nick, "webmaster") = 4 Or InStr(5, Nick, "webmaster") = 5 Or InStr(6, Nick, "webmaster") = 6 Or InStr(7, Nick, "webmaster") = 7 Or InStr(8, Nick, "webmaster") = 8 Or InStr(9, Nick, "webmaster") = 9 Then
a = a
ElseIf InStr(1, Nick, "ebmaster") = 1 Or InStr(1, Nick, "ebmaster") = 69 Or InStr(1, Nick, "ebmaster") = 70 Or InStr(1, Nick, "ebmaster") = 71 Or InStr(1, Nick, "ebmaster") = 71 Or InStr(1, Nick, "ebmaster") = 73 Or InStr(1, Nick, "ebmaster") = 74 Then
a = a
Else
List1.AddItem Nick
Names(y) = Nick
End If
Next y


'Text1.Text = Text1.Text & vbCrLf & BPData
'BPData = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.Text = Text1.Text & vbCrLf & "<Chat:#BPAdvCFG>(" & BPNick & ":) " & Text2.Text & vbCrLf
IRC.SendData "PRIVMSG " & BPChan & " :" & Text2.Text & vbCrLf
End If
End Sub

Private Sub Timer1_Timer()
Label2.Caption = IRC.State
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.tlb"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BPAdvCFG Tool by Nigel Todman"
   ClientHeight    =   4275
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   6930
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":482C2
   ScaleHeight     =   4275
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   900
      Left            =   0
      TabIndex        =   36
      Top             =   3400
      Width           =   6885
      Scrollbar       =   0   'False
      URL             =   "http://coinurl.com/get.php?id=35561"
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   1200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   2
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3960
      TabIndex        =   26
      Top             =   1920
      Width           =   175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3960
      TabIndex        =   24
      Top             =   1560
      Width           =   175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3960
      TabIndex        =   23
      Top             =   1200
      Width           =   175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1800
      TabIndex        =   22
      Top             =   2880
      Width           =   175
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Config.ini"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   3
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   2640
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   2
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   2520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   3
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   255
      Left            =   6240
      TabIndex        =   32
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   2
   End
   Begin MSComctlLib.Slider Slider6 
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   1
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raptr: veritas_"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   5520
      TabIndex        =   35
      Top             =   3240
      Width           =   1350
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BitCoin 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   240
      TabIndex        =   34
      Top             =   3240
      Width           =   3780
   End
   Begin VB.Image Image4 
      Height          =   4845
      Left            =   3960
      Picture         =   "Form1.frx":B2EB4
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Image Image3 
      Height          =   5685
      Left            =   720
      Picture         =   "Form1.frx":24EDDE
      Stretch         =   -1  'True
      Top             =   3600
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Image Image2 
      Height          =   5685
      Left            =   0
      Picture         =   "Form1.frx":3EAD08
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.Image Image1 
      Height          =   4605
      Left            =   0
      Picture         =   "Form1.frx":586C32
      Top             =   3960
      Width           =   7110
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Intro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2640
      TabIndex        =   27
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Music"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2640
      TabIndex        =   25
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Textures"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   21
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SSAO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   2640
      TabIndex        =   20
      Top             =   1200
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motion Blur"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   19
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Environment Map"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
      Width           =   1710
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   17
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OverallQuality"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HUDFullWidth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anti Aliasing"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contrast"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gamma"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aspect Ratio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOT Set"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   5430
      TabIndex        =   2
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Config.ini"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   930
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu launch 
         Caption         =   "Launch BurnoutParadise.exe"
      End
      Begin VB.Menu Save 
         Caption         =   "Save settings to Config.ini"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu TextC 
      Caption         =   "Text Color"
      Begin VB.Menu black 
         Caption         =   "Black"
      End
      Begin VB.Menu red 
         Caption         =   "Red"
      End
      Begin VB.Menu green 
         Caption         =   "Green"
      End
      Begin VB.Menu yellow 
         Caption         =   "Yellow"
      End
      Begin VB.Menu blue 
         Caption         =   "Blue"
      End
      Begin VB.Menu magenta 
         Caption         =   "Magenta"
      End
      Begin VB.Menu cyan 
         Caption         =   "Cyan"
      End
      Begin VB.Menu white 
         Caption         =   "White"
      End
   End
   Begin VB.Menu Background 
      Caption         =   "Background"
      Begin VB.Menu Burnout 
         Caption         =   "Burnout"
      End
      Begin VB.Menu bgblack 
         Caption         =   "Black"
      End
      Begin VB.Menu bgred 
         Caption         =   "Red"
      End
      Begin VB.Menu bgwhite 
         Caption         =   "White"
      End
   End
   Begin VB.Menu chat 
      Caption         =   "Chat"
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InstallPath, Build, NTVersion
Dim BPSettings
Dim BPConfig
Dim BPValues()
Dim BPLength
Dim BPTweak()
Dim ResTmp
Dim BPSound
Dim BPTelemetry
Dim ToolTipTmp
Dim MenuMusic
Dim NoIntro
Dim Chaturl As String
Dim x, y, Z, Tmp
Dim auto
Dim mclsToolTip As New clsToolTip
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub About_Click()
MsgBox ("BPAdvCFG " & Build & " by Nigel Todman" & vbCrLf & "Burnout Paradise Advanced Config Tool " & Build & vbCrLf & "E-Mail: nigel.todman@gmail.com" & vbCrLf & "Steam: Veritas_83" & vbCrLf & "Raptr: veritas_" & vbCrLf & "Twitter: @Veritas_83" & vbCrLf & "BitCoin: 18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ" & vbCrLf & "Web: nigeltodman.com")
End Sub

Private Sub bgblack_Click()
Form1.Picture = Image3.Picture
End Sub

Private Sub bgred_Click()
Form1.Picture = Image4.Picture
End Sub

Private Sub bgwhite_Click()
Form1.Picture = Image2.Picture
End Sub

Private Sub black_Click()
vbcolor = vbBlack
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub blue_Click()
vbcolor = vbBlue
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub Burnout_Click()
Form1.Picture = Image1.Picture
End Sub

Private Sub chat_Click()
'Randomize Timer
'Chaturl = "cbe004.chat.mibbit.com/?server=irc.WebChat.Org" & Chr(38) & "nick=BPGamer" & Chr(38) & "channel=%23BPAdvCFG"
'MsgBox Chaturl
'Shell ("cmd.exe /c start iexplore " & Chaturl)
'Shell ("cmd.exe /c start http://tinyurl.com/7jcjgt4")
'Shell ("cmd.exe /c start http://cbe004.chat.mibbit.com/?server=irc.WebChat.Org&nick=BPGamer&channel=%23BPAdvCFG")
'Shell ("cmd.exe /c start http://cbe004.chat.mibbit.com/?server=irc.WebChat.Org" & Chr(38) & "channel=%23BPAdvCFG")
'http://cbe004.chat.mibbit.com/?server=irc.WebChat.Org&channel=%23BPAdvCFG
Form2.Visible = True
End Sub

Private Sub Command1_Click()
If Len(BPSettings) < 4 Then
BPSettings = InputBox("Enter the full path to you Config.ini file. Usually in C:\Users\%USERNAME%\AppData\Local\Criterion Games\Burnout Paradise on Vista/Win7 C:\Documents and Settings\%USERNAME%\Local Settings\Application Data\Criterion Games\Burnout Paradise on Windows XP. Do not enter trailing \!")
Else
BPConfig = BPSettings + "\Config.ini"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(BPConfig) = True Then
    Label2.Caption = "Set."
    Open BPConfig For Input As #1
    Z = 0
    ReDim BPValues(420)
    Do While Not EOF(1)
        Z = Z + 1
        Line Input #1, BPValues(Z)
    Loop
    BPLength = Z
    Close #1
    ReDim BPValues(BPLength)
    Open BPConfig For Input As #1
    Z = 0
    Do While Not EOF(1)
        Z = Z + 1
        Line Input #1, BPValues(Z)
    Loop
    Close #1
    Open BPConfig For Input As #1
    For x = 1 To BPLength
        If Mid$(BPValues(x), 1, 6) = "Width=" Then
            Text1.Text = Mid$(BPValues(x), 7, Len(BPValues(x)) - 6)
        ElseIf Mid$(BPValues(x), 1, 7) = "Height=" Then
            Text2.Text = Mid$(BPValues(x), 8, Len(BPValues(x)) - 7)
        ElseIf Mid$(BPValues(x), 1, 12) = "AspectRatio=" Then
            Text3.Text = Mid$(BPValues(x), 13, Len(BPValues(x)) - 12)
        ElseIf Mid$(BPValues(x), 1, 16) = "GammaCorrection=" Then
            Text4.Text = Mid$(BPValues(x), 17, Len(BPValues(x)) - 16)
        ElseIf Mid$(BPValues(x), 1, 11) = "Brightness=" Then
            Text5.Text = Mid$(BPValues(x), 12, Len(BPValues(x)) - 11)
        ElseIf Mid$(BPValues(x), 1, 9) = "Contrast=" Then
            Text6.Text = Mid$(BPValues(x), 10, Len(BPValues(x)) - 9)
        'ElseIf Mid$(BPValues(x), 1, 13) = "AntiAliasing=" Then
        '    Text7.Text = Mid$(BPValues(x), 14, Len(BPValues(x)) - 13)
        ElseIf Mid$(BPValues(x), 1, 13) = "AntiAliasing=" Then
            'Slider4.Value = Mid$(BPValues(x), 14, Len(BPValues(x)) - 13)
            If Mid$(BPValues(x), 14, Len(BPValues(x)) - 13) = 0 Then
            Slider4.Value = 0
            ElseIf Mid$(BPValues(x), 14, Len(BPValues(x)) - 13) = 2 Then
            Slider4.Value = 1
            ElseIf Mid$(BPValues(x), 14, Len(BPValues(x)) - 13) = 4 Then
            Slider4.Value = 2
            ElseIf Mid$(BPValues(x), 14, Len(BPValues(x)) - 13) = 8 Then
            Slider4.Value = 3
            End If
        ElseIf Mid$(BPValues(x), 1, 14) = "HUDFullWidth=0" Then
            Check1.Value = 0
        ElseIf Mid$(BPValues(x), 1, 14) = "HUDFullWidth=1" Then
            Check1.Value = 1
        'ElseIf Mid$(BPValues(x), 1, 15) = "OverallQuality=" Then
        '    Text9.Text = Mid$(BPValues(x), 16, Len(BPValues(x)) - 15)
        ElseIf Mid$(BPValues(x), 1, 15) = "OverallQuality=" Then
            Slider2.Value = Mid$(BPValues(x), 16, Len(BPValues(x)) - 15)
        'ElseIf Mid$(BPValues(x), 1, 8) = "Shadows=" Then
        '    Text10.Text = Mid$(BPValues(x), 9, Len(BPValues(x)) - 8)
        ElseIf Mid$(BPValues(x), 1, 8) = "Shadows=" Then
            Slider5.Value = Mid$(BPValues(x), 9, Len(BPValues(x)) - 8)
        'ElseIf Mid$(BPValues(x), 1, 15) = "EnvironmentMap=" Then
        '    Text11.Text = Mid$(BPValues(x), 16, Len(BPValues(x)) - 15)
        ElseIf Mid$(BPValues(x), 1, 15) = "EnvironmentMap=" Then
            Slider6.Value = Mid$(BPValues(x), 16, Len(BPValues(x)) - 15)
        'ElseIf Mid$(BPValues(x), 1, 11) = "MotionBlur=" Then
        '    Text12.Text = Mid$(BPValues(x), 12, Len(BPValues(x)) - 11)
        ElseIf Mid$(BPValues(x), 1, 11) = "MotionBlur=" Then
            Slider3.Value = Mid$(BPValues(x), 12, Len(BPValues(x)) - 11)
        ElseIf Mid$(BPValues(x), 1, 6) = "SSAO=0" Then
            Check3.Value = 0
        ElseIf Mid$(BPValues(x), 1, 12) = "SSAO=1" Then
            Check3.Value = 1
        'ElseIf Mid$(BPValues(x), 1, 9) = "Textures=" Then
        '    Text8.Text = Mid$(BPValues(x), 10, Len(BPValues(x)) - 9)
        ElseIf Mid$(BPValues(x), 1, 9) = "Textures=" Then
            Slider1.Value = Mid$(BPValues(x), 10, Len(BPValues(x)) - 9)
        ElseIf Mid$(BPValues(x), 1, 11) = "[Sound]" Then
            BPSound = x
        ElseIf Mid$(BPValues(x), 1, 11) = "[Telemetry]" Then
            BPTelemetry = x
        End If
    Next x
    Close #1
    MenuMusic = InstallPath + "\SOUND\STREAMS\GUNS_AND_ROSES.SNS"
    If fso.FileExists(MenuMusic) = False Then
        Shell ("cmd.exe /c echo/ >> " & Chr(34) & MenuMusic & Chr(34))
    Else
        If FileLen(MenuMusic) > 24 Then
            Check2.Value = 1
        ElseIf FileLen(MenuMusic) < 24 Then
            Check2.Value = 0
        End If
    End If
    If FileLen(InstallPath + "\VIDEOS\CRITERION.VP6") < 70 Then
        Check4.Value = 1
    ElseIf FileLen(InstallPath + "\VIDEOS\CRITERION.VP6") > 70 Then
        Check4.Value = 0
    End If
    'MsgBox BPValues(1) & vbCrLf & BPValues(2) & vbCrLf & BPValues(3) & vbCrLf & BPValues(4) & vbCrLf & BPValues(5)
End If
End If
End Sub

Private Sub cyan_Click()
vbcolor = vbCyan
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Public Function hex2ascii(ByVal hextext As String) As String
For y = 1 To Len(hextext)
    num = Mid(hextext, y, 2)
    Value = Value & Chr(Val("&h" & num))
    y = y + 1
Next y
hex2ascii = Value
End Function
Private Sub Form_Load()
On Error Resume Next
auto = MsgBox("Would you like to try and automatically detect settings path?", vbYesNo)
'Form2.Visible = False
'Form2.Enabled = True
'http://coinurl.com/get.php?id=35561
'Scriptlet1.Width = 3650
'Scriptlet1.Height = 905
Scriptlet1.Width = 7025
Scriptlet1.Height = 905
Build = "v0.5"
'MsgBox "Build String set to: " & Build
Form1.Caption = "BPAdvCFG Tool " & Build & " (Final Release)"
'MsgBox "Titlebar set"
NTVersion = GetStringValue("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion", "CurrentVersion")
'MsgBox "NTVersion set, removing null characters.."
NTVersion = Left$(NTVersion, InStr(NTVersion, Chr$(0)) - 1)
'MsgBox "NTVersion result: " & NTVersion
InstallPath = GetStringValue("HKEY_LOCAL_MACHINE\Software\EA Games\Burnout(TM) Paradise The Ultimate Box", "Install Dir")
'MsgBox "Detecting install path..."
InstallPath = Left$(InstallPath, InStr(InstallPath, Chr$(0)) - 1)
'MsgBox "InstallPath result: " & InstallPath
'BPSettings = "C:\Users\Administrator\AppData\Local\Criterion Games\Burnout Paradise"

Text7.Visible = False
Text8.Visible = False
Text9.Visible = False
Text10.Visible = False
Text11.Visible = False
Text12.Visible = False

'ToolTip mod
Dim ctrl As Control
  With mclsToolTip
    '
    ' Create the tooltip window.
    '
    Call .Create(Me)

    '
    ' Set the tooltip's width so that it displays
    ' multiline text and no tool's line length exceeds
    ' roughly 240 pixels.
    '
    .MaxTipWidth = 240

    '
    ' Show the tooltip for 20 seconds.
    '
    .DelayTime(ttDelayShow) = 20000
    
    '
    ' Add a tooltip tool to each control on the Form.
    '
    For Each ctrl In Controls
        Call .AddTool(ctrl)
    Next

ToolTipTmp = "The display resolution of a display device is the number of distinct pixels in each dimension that can be displayed. " & vbCrLf
ToolTipTmp = ToolTipTmp + "It can be an ambiguous term especially as the displayed resolution is controlled by all different factors in cathode ray " & vbCrLf
ToolTipTmp = ToolTipTmp + "tube (CRT), flat panel or projection displays using fixed picture-element (pixel) arrays." & vbCrLf
ToolTipTmp = ToolTipTmp + "It is usually quoted as width × height, with the units in pixels: for example, " & Chr(34) & "1024×768" & Chr(34) & " means the width is 1024 pixels" & vbCrLf
ToolTipTmp = ToolTipTmp + "and the height is 768 pixels. This example would normally be spoken as " & Chr(34) & "ten twenty-four by seven sixty-eight" & Chr(34) & vbCrLf & " or " & Chr(34) & "ten twenty-four by seven six eight" & Chr(34)
.ToolText(Text1) = ToolTipTmp
ToolTipTmp = "Valid Values are 4x3, 16x9, 5x4, 16x10, 15x9, 15x10, Auto (Default Auto)"
.ToolText(Text3) = ToolTipTmp
ToolTipTmp = "Min 0.000 Max 3.000 (Default 1.000)"
.ToolText(Text4) = ToolTipTmp
ToolTipTmp = "Min 1 Max 100 (Default 50)"
.ToolText(Text5) = ToolTipTmp
ToolTipTmp = "Min 1 Max 100 (Default 50)"
.ToolText(Text6) = ToolTipTmp
ToolTipTmp = "Super sampling anti-aliasing (SSAA), also called full-scene anti-aliasing (FSAA), is used to avoid aliasing (or 'jaggies')" & vbCrLf
ToolTipTmp = ToolTipTmp + "on full-screen images. SSAA was the first type of anti-aliasing available with early video cards. But due to its tremendous computational" & vbCrLf
ToolTipTmp = ToolTipTmp + "cost and with the advent of multisample anti-aliasing (MSAA) support on GPUs, it is no longer widely used in real time applications. MSAA" & vbCrLf
ToolTipTmp = ToolTipTmp + "provides somewhat lower graphic quality, but also tremendous savings in computational power" & vbCrLf & vbCrLf
ToolTipTmp = ToolTipTmp + "Full-scene anti-aliasing by supersampling usually means that each full frame is rendered at double (2x) or quadruple (4x) the display resolution," & vbCrLf
ToolTipTmp = ToolTipTmp + "and then down-sampled to match the display resolution. So a 2x FSAA would render 4 supersampled pixels for each single pixel of each frame. While" & vbCrLf
ToolTipTmp = ToolTipTmp + "rendering at larger resolutions will produce better results, more processor power is needed which can degrade performance and frame rate." & vbCrLf
ToolTipTmp = ToolTipTmp + "(Default 0, Valid values are 0,2,4,8)"
'.ToolText(Text7) = ToolTipTmp
.ToolText(Slider4) = ToolTipTmp
ToolTipTmp = "0 = Off 1 = Bumper Cam Only 2 = Always On (Default 2)"
'.ToolText(Text12) = ToolTipTmp
.ToolText(Slider3) = ToolTipTmp
ToolTipTmp = "0 = Standard 1 = High 2 = Highest (Default 0)"
'.ToolText(Text8) = ToolTipTmp
.ToolText(Slider1) = ToolTipTmp
ToolTipTmp = "0 = Standard 1 = High 2 = Highest 3 = Custom (Default 0, Recommended set this to 3)"
'.ToolText(Text9) = ToolTipTmp
.ToolText(Slider2) = ToolTipTmp
ToolTipTmp = "0 = Low 1 = Medium 2 = High (Default 1)"
'.ToolText(Text10) = ToolTipTmp
.ToolText(Slider5) = ToolTipTmp
ToolTipTmp = "Environment mapping is a form of texture mapping in which the texture coordinates are view-dependent. One common application, for example, is to simulate reflection" & vbCrLf
ToolTipTmp = ToolTipTmp + "on a shiny object. One can environment map the interior of a room to a metal cup in a room. As the viewer moves about the cup, the texture coordinates of the cup’s vertices move" & vbCrLf
ToolTipTmp = ToolTipTmp + "accordingly, providing the illusion of reflective metal. (Default 0, Valid values are 0 = Low 1 = Standard)"
'.ToolText(Text11) = ToolTipTmp
.ToolText(Slider6) = ToolTipTmp
ToolTipTmp = "Overwrites Menu Music file GUNS_AND_ROSES.SNS with a much quieter 3 byte version, Backups original to *.SN0, Rechecking and saving will restore the backup."
.ToolText(Check2) = ToolTipTmp
ToolTipTmp = "Screen Space Ambient Occlusion (SSAO) is a rendering technique for efficiently approximating the well-known computer graphics ambient occlusion effect in real time." & vbCrLf
ToolTipTmp = ToolTipTmp + "It was developed by Vladimir Kajalin while working at Crytek and was used for the first time in a video game in the 2007 Windows game Crysis made by Crytek." & vbCrLf & vbCrLf
ToolTipTmp = ToolTipTmp + "The algorithm is implemented as a pixel shader, analyzing the scene depth buffer which is stored in a texture. For every pixel on the screen, the pixel shader samples the depth" & vbCrLf
ToolTipTmp = ToolTipTmp + "values around the current pixel and tries to compute the amount of occlusion from each of the sampled points. In its simplest implementation, the occlusion factor depends only" & vbCrLf
ToolTipTmp = ToolTipTmp + "on the depth difference between sampled point and current point. (Default Off)"
.ToolText(Check3) = ToolTipTmp
ToolTipTmp = "Overwrites Intro logo videos (EA, Criterion) files with much shorter 68 byte versions, Backups originals to *.VP0, Unchecking and saving will restore the backup."
.ToolText(Check4) = ToolTipTmp

End With

'EnvironmentMap Max Standard Min Low
'Gamma Max 3.000 Min 0.000 Default 1.000
'AspectRatio 4x3 16x9 5x4 16x10 15x9 15x10 Auto
'MotionBlur 0 = off 1 = Bumper Cam Only 2 = Always On
'Shadows Low, Medium, High
'Textures Standard =0 High =1 Highest =2
'OverallQuality 0 = standard 1 = high 2 = highest 3 = custom

'No Menu Music Fix!
'InstallPath + "\SOUND\STREAMS\GUNS_AND_ROSES.SNS" Backup and patch to 0 byte!

'No Intro Fix
'InstallPath + "\VIDEOS\CRITERION.VP6"

'4D 56 68 64 20 00 00 00 76 70 36 30 20 00 10 00 01 00 00 00 1A 00 00 00 BE 7B 00 00 21 04 00 00 4D 56 30 4B 22 00 00 00 29 30 00 0D 01 02 01 02 00 00 00 00 00 FE 14 A8 81 ED 7D FE F8 B8 00 00 00 00
'MsgBox "Declaring NoIntro data..."
NoIntro = "4D566864200000007670363020001000010000001A000000BE7B0000210400004D56304B220000002930000D010201020000000000FE14A881ED7DFEF8B800000000"
'MsgBox "Converting NoIntro data from Hex to ASCII"
NoIntro = hex2ascii(NoIntro)
'MsgBox "Enumerating Username..."

If auto = 6 Then
Shell ("cmd.exe /c echo %USERNAME% >> user.temp"), vbHide
Sleep (100)
Open VB.App.Path + "\user.temp" For Input As #420
Line Input #420, Tmp
Tmp = Left$(Tmp, InStr(Tmp, " ") - 1)
CurrUser = Tmp
Close #420
Sleep (50)
'MsgBox "Username enumerated as: " & CurrUser & " Validating..."

If NTVersion = "5.1" Then
    'launch.Enabled = False
    If fso.FileExists("C:\Documents and Settings\" & CurrUser & "\Local Settings\Application Data\Criterion Games\Burnout Paradise\Config.ini") = True Then
        BPSettings = "C:\Documents and Settings\" & CurrUser & "\Local Settings\Application Data\Criterion Games\Burnout Paradise"
        MsgBox "Settings path auto-detected :) Click Set Config.ini to detect values."
    End If
Else
    If fso.FileExists("C:\Users\" & CurrUser & "\AppData\Local\Criterion Games\Burnout Paradise\Config.ini") = True Then
        BPSettings = "C:\Users\" & CurrUser & "\AppData\Local\Criterion Games\Burnout Paradise"
        MsgBox "Settings path auto-detected :) Click Set Config.ini to detect values."
    Else
    BPSettings = vbNull
    End If
End If

ElseIf auto = 7 Then
BPSettings = vbNull
End If

'MsgBox "Username validated...deleting temp file"
Shell ("cmd.exe /c del user.temp"), vbHide


'Red = 255


'InstallPath + "\VIDEOS\CRITERION.VP6"
'InstallPath + "\VIDEOS\EAFRANCHISE.VP6"
'InstallPath + "\VIDEOS\EAHD.VP6"
'InstallPath + "\SOUND\STREAMS\EAFRANCHISE.SNS"
'InstallPath + "\SOUND\STREAMS\CRITERION.SNS"
'MsgBox "Initalization complete :)"
End Sub

Private Sub green_Click()
vbcolor = vbGreen
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub


Private Sub Label18_Click()
Clipboard.Clear
Clipboard.SetText "18j2Env7QokhGG5MccS3LPBKnjsko6u4NQ", vbCFText
MsgBox "BitCoin address now in clipboard! Ctrl+V to paste."
End Sub

Private Sub Label19_Click()
Shell ("cmd.exe /c start http://raptr.com/veritas_")
End Sub

Private Sub Launch_Click()
If NTVersion = "5.1" Then
Shell ("cmd.exe /c " & Chr(34) & InstallPath & "BurnoutLauncher.exe" & Chr(34))
Else
Shell ("cmd.exe /c " & Chr(34) & InstallPath & "BurnoutLauncher.exe" & Chr(34))
End If
End Sub

Private Sub launch2_Click()
Shell ("cmd.exe /c " & Chr(34) & InstallPath & "BurnoutLauncher.exe" & Chr(34) & " -multithread")
End Sub

Private Sub magenta_Click()
vbcolor = vbMagenta
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub red_Click()
vbcolor = vbRed
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub Save_Click()
'Shell ("cmd.exe /c start http://nigelt.wordpress.com")
If Label2.Caption = "Set." Then
MsgBox "Creating Backup of Current Settings to Config.in0", vbExclamation
Open BPSettings + "\Config.in0" For Output As #2
For x = 1 To BPLength
Print #2, BPValues(x)
Next x
Print #2, "!Backup File Generated by BPAdvCFG " & Build & " - Updates? http://www.nigeltodman.com"
Close #2
End If

ReDim BPTweak(BPLength + 24)
BPTweak(1) = "[Display]"
BPTweak(2) = "AdapterIndex=0"
BPTweak(3) = "Width=" & Text1.Text
BPTweak(4) = "Height=" & Text2.Text
BPTweak(5) = "AspectRatio=" & Text3.Text
BPTweak(6) = "AdapterName="
BPTweak(7) = "GammaCorrection=" & Text4.Text
BPTweak(8) = "Brightness=" & Text5.Text
BPTweak(9) = "Contrast=" & Text6.Text
BPTweak(10) = "[Settings]"
'BPTweak(11) = "AntiAliasing=" & Text7.Text
'BPTweak(11) = "AntiAliasing=" & Slider4.Value
If Slider4.Value = 0 Then
BPTweak(11) = "AntiAliasing=0"
ElseIf Slider4.Value = 1 Then
BPTweak(11) = "AntiAliasing=2"
ElseIf Slider4.Value = 2 Then
BPTweak(11) = "AntiAliasing=4"
ElseIf Slider4.Value = 3 Then
BPTweak(11) = "AntiAliasing=8"
End If
BPTweak(12) = "NumMonitors=1"
BPTweak(13) = "CarMonitor=1"
If Check1.Value = 1 Then
BPTweak(14) = "HUDFullWidth=1"
ElseIf Check1.Value = 0 Then
BPTweak(14) = "HUDFullWidth=0"
End If
'BPTweak(15) = "OverallQuality=" & Text9.Text
BPTweak(15) = "OverallQuality=" & Slider2.Value
BPTweak(16) = "Webcam="
'BPTweak(17) = "Shadows=" & Text10.Text
BPTweak(17) = "Shadows=" & Slider5.Value
'BPTweak(18) = "EnvironmentMap=" & Text11.Text
BPTweak(18) = "EnvironmentMap=" & Slider6.Value
'BPTweak(19) = "MotionBlur=" & Text12.Text
BPTweak(19) = "MotionBlur=" & Slider3.Value
If Check3.Value = 1 Then
BPTweak(20) = "SSAO=1"
ElseIf Check3.Value = 0 Then
BPTweak(20) = "SSAO=0"
End If

'BPTweak(21) = "Textures=" & Text8.Text
BPTweak(21) = "Textures=" & Slider1.Value

If Len(BPSound) > 1 Then
    For x = BPSound To BPLength
        BPTweak(x) = BPValues(x)
    Next x
ElseIf Len(BPTelemetry) > 1 Then
    For x = BPTelemetry To BPLength
        BPTweak(x) = BPValues(x)
    Next x
Else
    For x = 22 To BPLength
        BPTweak(x) = BPValues(x)
    Next x
End If

MenuMusic = InstallPath + "SOUND\STREAMS\GUNS_AND_ROSES.SNS"


If Check2.Value = 1 Then
    If FileLen(MenuMusic) < 4 Then
    Shell ("cmd.exe /c copy " & Chr(34) & InstallPath + "SOUND\STREAMS\GUNS_AND_ROSES.SN0" & Chr(34) & " " & Chr(34) & MenuMusic & Chr(34))
    End If
ElseIf Check2.Value = 0 Then
    If FileLen(MenuMusic) > 4 Then
    Shell ("cmd.exe /c move " & Chr(34) & MenuMusic & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\GUNS_AND_ROSES.SN0" & Chr(34))
    Sleep (2000)
    Shell ("cmd.exe /c echo/ >> " & Chr(34) & MenuMusic & Chr(34))
    End If
End If

If Check4.Value = 1 Then
    If FileLen(InstallPath + "VIDEOS\CRITERION.VP6") > 70 Then
    Open VB.App.Path & "\NoIntro.bat" For Output As #4
    Print #4, "move " & Chr(34) & InstallPath + "VIDEOS\CRITERION.VP6" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\CRITERION.VP0" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "VIDEOS\EAFRANCHISE.VP6" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\EAFRANCHISE.VP0" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "VIDEOS\EAHD.VP6" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\EAHD.VP0" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SNS" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SN0" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SNS" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SN0" & Chr(34)
'    Print #4, "del " & Chr(34) & InstallPath + "VIDEOS\CRITERION.VP6" & Chr(34)
'    Print #4, "del " & Chr(34) & InstallPath + "VIDEOS\EAFRANCHISE.VP6" & Chr(34)
'    Print #4, "del " & Chr(34) & InstallPath + "VIDEOS\EAHD.VP6" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SNS" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SN0" & Chr(34)
    Print #4, "move " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SNS" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SN0" & Chr(34)
    Close #4
    Shell (VB.App.Path & "\NoIntro.bat")
    Sleep (2000)
    Open InstallPath + "\VIDEOS\CRITERION.VP6" For Output As #5
    Print #5, NoIntro
    Close #5
    Open InstallPath + "\VIDEOS\EAFRANCHISE.VP6" For Output As #6
    Print #6, NoIntro
    Close #6
    Open InstallPath + "\VIDEOS\EAHD.VP6" For Output As #7
    Print #7, NoIntro
    Close #7
    End If
ElseIf Check4.Value = 0 Then
    If FileLen(InstallPath + "VIDEOS\CRITERION.VP6") < 70 Then
    Open VB.App.Path & "\ReIntro.bat" For Output As #8
    Print #8, "copy " & Chr(34) & InstallPath + "VIDEOS\CRITERION.VP0" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\CRITERION.VP6" & Chr(34)
    Print #8, "copy " & Chr(34) & InstallPath + "VIDEOS\EAFRANCHISE.VP0" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\EAFRANCHISE.VP6" & Chr(34)
    Print #8, "copy " & Chr(34) & InstallPath + "VIDEOS\EAHD.VP0" & Chr(34) & " " & Chr(34) & InstallPath + "VIDEOS\EAHD.VP6" & Chr(34)
    Print #8, "copy " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SN0" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\CRITERION.SNS" & Chr(34)
    Print #8, "copy " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SN0" & Chr(34) & " " & Chr(34) & InstallPath + "SOUND\STREAMS\EAFRANCHISE.SNS" & Chr(34)
    Close #8
    Shell (VB.App.Path & "\ReIntro.bat")
    End If
End If
    
'InstallPath + "\SOUND\STREAMS\EAFRANCHISE.SNS
Shell ("cmd.exe /c attrib -r " & BPSettings + "\Config.ini"), vbHide
Shell ("cmd.exe /c attrib -r " & BPSettings + "\Config.ini"), vbHide
MsgBox "Config File Generated. Press OK to write file Config.ini!", vbInformation
Open BPSettings + "\Config.ini" For Output As #3
For x = 1 To BPLength
Print #3, BPTweak(x)
Next x
Close #3
MsgBox ("Settings written to Config.ini!"), vbInformation
'Shell ("cmd.exe /c attrib +r " & BPSettings + "\Config.ini"), vbHide
Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\ReIntro.bat" & Chr(34))
Shell ("cmd.exe /c del " & Chr(34) & VB.App.Path & "\NoIntro.bat" & Chr(34))
End Sub

Private Sub white_Click()
vbcolor = vbWhite
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

Private Sub yellow_Click()
vbcolor = vbYellow
Label1.ForeColor = vbcolor
Label2.ForeColor = vbcolor
Label3.ForeColor = vbcolor
Label4.ForeColor = vbcolor
Label5.ForeColor = vbcolor
Label6.ForeColor = vbcolor
Label7.ForeColor = vbcolor
Label8.ForeColor = vbcolor
Label9.ForeColor = vbcolor
Label10.ForeColor = vbcolor
Label11.ForeColor = vbcolor
Label12.ForeColor = vbcolor
Label13.ForeColor = vbcolor
Label14.ForeColor = vbcolor
Label15.ForeColor = vbcolor
Label16.ForeColor = vbcolor
Label17.ForeColor = vbcolor
Label18.ForeColor = vbcolor
Label19.ForeColor = vbcolor
End Sub

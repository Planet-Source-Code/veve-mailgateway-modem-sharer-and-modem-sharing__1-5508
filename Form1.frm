VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MGv1"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4215
      Begin RichTextLib.RichTextBox Rtb 
         Height          =   2535
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4471
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0442
      End
      Begin VB.Label Label1 
         Caption         =   "Mail Gateway Version 1.0"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":0516
         Top             =   120
         Width           =   480
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2040
      Top             =   3960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock LocalPOP3 
      Index           =   0
      Left            =   1080
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   110
   End
   Begin MSWinsockLib.Winsock LocalSMTP 
      Index           =   0
      Left            =   120
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   25
   End
   Begin MSWinsockLib.Winsock RemoteSMTP 
      Index           =   0
      Left            =   600
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "bow.intnet.mu"
      RemotePort      =   25
   End
   Begin MSWinsockLib.Winsock RemotePOP3 
      Index           =   0
      Left            =   1560
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "bow.intnet.mu"
      RemotePort      =   110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Mail Gateway Version 1.0
'Written by V. Veerapen
'
'Background
'
'This code allows several users on a LAN to access their mail
'from an ISP by using Windows NT 4.0 RAS services. Only ONE modem is
'needed!!!
'

Option Explicit

Const Occupied = 1
Const Vacant = 0

Dim SessionTable(255) As Long
Dim SessionCount As Long

Private Sub Form_Activate()
  
  On Error GoTo SocketError
  LocalSMTP(0).Listen
  LocalPOP3(0).Listen
  On Error GoTo 0
  Say "Listening on ports 25 and 110"
  
  Exit Sub
  
SocketError:
  Say "WARNING ! - Ports 25 and/or 110 already in use"
  
End Sub
'Check if machine is connected to the Internet
Private Function IsOnline() As Boolean

  IsOnline = InternetGetConnectedState(0&, 0&)
  
End Function
'Check if the RAS window is already open
Private Function IsConnecting() As Boolean

  IsConnecting = FindWindow(vbNullString, "Dial-Up Networking")

End Function
'Connect to the Internet via RAS
Private Sub ConnectToNet()

  Dim Result
  
  Say "Connecting to the Internet"
  
  Result = Shell("rasphone", 1)
  SendKeys "{enter}", True

  DoEvents

End Sub
'Disconnect from the Internet
Private Sub DisconnectFromNet()

  Dim Result
  
  Say "Disconnecting from the Internet"
  
  Result = Shell("rasphone -h " & "Internet", 1)
  DoEvents  '                          ^--- Phone book entry to hangup goes here
  
End Sub
'Get a free socket
Private Function GetFreeSocket() As Long

  Dim i As Long
    
  For i = 1 To 255
    If SessionTable(i) = 0 Then
      GetFreeSocket = i
      Exit Function
    End If
  Next

  Say "WARNING ! - Maximum session limit exceeded !"
  GetFreeSocket = -1    'No free sockets!

End Function

Private Sub LocalPOP3_ConnectionRequest(Index As Integer, ByVal requestID As Long)

  Dim i As Long
  
  i = GetFreeSocket()
  If i <> -1 Then
    
    Timer1.Enabled = False
    
    SessionTable(i) = Occupied
    SessionCount = SessionCount + 1
    
    Load LocalPOP3(i)
    Load RemotePOP3(i)
    
    If IsConnecting = False And IsOnline = False Then ConnectToNet
    Do While IsOnline = False
      DoEvents
    Loop

    LocalPOP3(i).Accept requestID
    RemotePOP3(i).Connect
        
  End If

End Sub

Private Sub LocalSMTP_ConnectionRequest(Index As Integer, ByVal requestID As Long)

  Dim i As Long
  
  i = GetFreeSocket
  If i <> -1 Then
  
    Timer1.Enabled = False
  
    SessionTable(i) = Occupied
    SessionCount = SessionCount + 1
  
    Load LocalSMTP(i)
    Load RemoteSMTP(i)
  
    If IsConnecting = False And IsOnline = False Then ConnectToNet
    Do While IsOnline = False
      DoEvents
    Loop
    
    LocalSMTP(i).Accept requestID
    RemoteSMTP(i).Connect
            
  End If

End Sub

'Transfer data from local POP3 client to remote server
Private Sub LocalPOP3_DataArrival(Index As Integer, ByVal bytesTotal As Long)

  Dim strData As String
  LocalPOP3(Index).GetData strData
  RemotePOP3(Index).SendData strData

End Sub

'Transfer data from POP3 server to local client
Private Sub RemotePOP3_DataArrival(Index As Integer, ByVal bytesTotal As Long)

  Dim strData As String
  RemotePOP3(Index).GetData strData
  LocalPOP3(Index).SendData strData

End Sub

'Transfer data from local SMTP client to remote server
Private Sub LocalSMTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)

  Dim strData As String
  LocalSMTP(Index).GetData strData
  RemoteSMTP(Index).SendData strData

End Sub

'Transfer data from SMTP server to local client
Private Sub RemoteSMTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)

  Dim strData As String
  RemoteSMTP(Index).GetData strData
  LocalSMTP(Index).SendData strData

End Sub

'Free up the local and remote POP3 socket
Private Sub LocalPOP3_Close(Index As Integer)

  LocalPOP3(Index).Close
  RemotePOP3(Index).Close
  
  Unload LocalPOP3(Index)
  Unload RemotePOP3(Index)
  
  SessionTable(Index) = 0
  SessionCount = SessionCount - 1
  
  If SessionCount = 0 Then Timer1.Enabled = True

End Sub

'Free up the local and remote SMTP socket
Private Sub LocalSMTP_Close(Index As Integer)

  LocalSMTP(Index).Close
  RemoteSMTP(Index).Close
  
  Unload LocalSMTP(Index)
  Unload RemoteSMTP(Index)
  
  SessionTable(Index) = 0
  SessionCount = SessionCount - 1

  If SessionCount = 0 Then Timer1.Enabled = True

End Sub

'Wait a bit before hanging up the connection
Private Sub Timer1_Timer()

  Timer1.Enabled = False
  If SessionCount = 0 Then DisconnectFromNet

End Sub

'Output to text box
Private Sub Say(sText As String)

  Rtb.Text = sText & vbCrLf & Rtb.Text

End Sub

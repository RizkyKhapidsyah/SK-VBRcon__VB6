VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmRcon 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Rcon"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmRcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Text            =   "27015"
      Top             =   2700
      Width           =   1395
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   2700
      Width           =   1635
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   315
      Left            =   6000
      TabIndex        =   7
      Top             =   2820
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "connect"
      Height          =   315
      Left            =   6000
      TabIndex        =   6
      Top             =   2460
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Top             =   2700
      Width           =   2475
   End
   Begin RichTextLib.RichTextBox txtGet 
      Height          =   2355
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4154
      _Version        =   393217
      BackColor       =   8421504
      ScrollBars      =   3
      TextRTF         =   $"frmRcon.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   3300
      Width           =   795
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   3300
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   7260
      Top             =   3180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Server Port"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   2460
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Server IP"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2460
      Width           =   2595
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Rcon Command"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   3060
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Rcon Password"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   3300
      TabIndex        =   4
      Top             =   2460
      Width           =   2475
   End
End
Attribute VB_Name = "frmRcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAXCHAR = 64000
Const ByteCode = "ÿÿÿÿ"
Dim delim As Byte
Dim rconNumber As String
Dim rconCommand As String
Dim passedCheck As Boolean
Dim sentChallenge As Boolean

Private Sub cmdConnect_Click()
    Dim sPort As Long
    sPort = CInt(txtPort.Text)
    wsk.RemoteHost = txtIP.Text ' The IP you wish to connect to
    wsk.RemotePort = sPort ' The port you wish to connect to
    wsk.Bind 554 ' Reserve 554
    sendChallenge
    sentChallenge = True
    cmdSend.Enabled = True
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
    wsk.Close
    cmdSend.Enabled = False
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub cmdSend_Click()
    sendCommand
End Sub

Private Sub Form_Load()
   cmdSend.Enabled = False
   cmdConnect.Enabled = True
   cmdDisconnect.Enabled = False
   sentChallenge = False
End Sub
Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
   Dim strReceived As String
   wsk.GetData strReceived
   checkChallenge strReceived
   If Len(strReceived) = 7 Then
        strReceived = "Command Successful..." & vbCrLf
   End If
   LogText txtGet, CStr(Replace(strReceived, ByteCode, "", 1)), vbBlack
   'txtGet.Text = strReceived
End Sub
Public Sub LogText(rtfBox As RichTextBox, strData As String, strColor As String)
    'Sub logtext to given richtextbox with hex color
        Dim strTemp As String
        Dim intRed As Integer, intGreen As Integer, intBlue As Integer
        
        strTemp = strColor
        'Parse the red, green and blue back to rgb
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        With rtfBox
            If Len(strData) + Len(.Text) > MAXCHAR Then
            'Scroll some text off the top to make more room
            .Text = Mid$(.Text, InStr(100 + Len(strData), .Text, vbCrLf) + 2)
            End If
            .SelStart = Len(.Text)
            .SelColor = RGB(intRed, intGreen, intBlue)
            Call DoColor(rtfBox, strData)
            .SelStart = Len(.Text)
        End With
End Sub
Function sendCommand()
    Dim myarray
    If Not Len(txtSend.Text) = 0 Then
        myarray = Split(txtSend.Text, " ", -1)
        Select Case myarray(0)
            Case "status"
                '"*(.+) (\d+) (\d+) *(-?\d+) *(\d?\d?:?\d\d:\d\d) *(\d+) *(\d+?) *(.+?)$"
                'This is a regular expression to parse players from status
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & " status"
                End If
            Case "changelevel"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "kick"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "banid"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "users"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "user"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "maps"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case "say"
                If sentChallenge Then
                    wsk.SendData ByteCode & "rcon " & rconNumber & " " & Chr(34) & txtPassword.Text & Chr(34) & " " & txtSend.Text
                End If
            Case Else
                wsk.SendData ByteCode & txtSend.Text
        End Select
    End If
    End Function
Function sendChallenge()
    wsk.SendData ByteCode & "challenge rcon"
    DoEvents
End Function
Function checkChallenge(strData As String)
    If Left(strData, 13) = "ÿÿÿÿchallenge" Then
        myarray = Split(strData, " ", -1)
        rconNumber = Left(myarray(2), Len(myarray(2)) - 2)
    End If
End Function

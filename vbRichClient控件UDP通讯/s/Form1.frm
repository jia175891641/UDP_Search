VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7200
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "测试"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AT+WSCAN"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AT+UART"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "搜索"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents udp1 As cUDP
Attribute udp1.VB_VarHelpID = -1
Dim ipp As String
Dim ME_IP As String

Private Sub Command1_Click() '通过udp发送给 主机127.0.0.1:5618
Text2.Text = ""
udp_send "www.usr.cn", "255,255,255,255"
End Sub


Private Sub udp_send(s As String, ip As String)
Dim dd() As Byte
 dd() = StrConv(s, vbFromUnicode)
 udp1.RemoteIP = "255.255.255.255"
 udp1.RemotePort = 20108
 udp1.SendData VarPtr(dd(0)), UBound(dd) + 1
End Sub




 
Private Sub Command2_Click(Index As Integer)
udp_send Command2(Index).Caption & vbCrLf, "10,10,100,254"
End Sub

Private Sub Form_Load()
Set udp1 = New cUDP

ME_IP = udp1.GetIP(VBA.Environ("computername"))
Me.Caption = "本机IP:" & ME_IP
udp1.Bind ME_IP, 488

End Sub


Private Sub Text2_DblClick()
Text2.Text = ""
End Sub

Private Sub udp1_NewDatagram(ByVal BytesTotal As Long, ByVal FirstBufferAfterOverflow As Boolean)
Dim d() As Byte, s$
ReDim d(BytesTotal - 1)
udp1.GetData VarPtr(d(0)), BytesTotal
Text2.Text = Text2.Text + StrConv(d, vbUnicode) + vbCrLf
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinSPY"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   11910
   Begin VB.Frame Frame4 
      Caption         =   "Other Settings"
      Height          =   1935
      Left            =   9400
      TabIndex        =   15
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox Check1 
         Caption         =   "Screen ReCapture : "
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   7
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         Top             =   1500
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Mouse Move Length, for Popup Menu : Move"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Screen Receive"
      Height          =   1935
      Left            =   5520
      TabIndex        =   11
      Top             =   120
      Width           =   3780
      Begin MSComctlLib.ProgressBar FileBar 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbBytesReceived 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bytes Received"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3555
      End
      Begin VB.Label lbFilereceived 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "File Received"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Capture Senttings"
      Height          =   1935
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CAPTURE"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Screen Resulation : "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1320
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "127.0.0.1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "1453"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   285
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Remote Port :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Remote IP :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   20
      Top             =   1950
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Osman Gulercan, Istanbul, TURKEY
'osmangulercan@yahoo.com
'o.gulercan@veezy.com

Dim Dosya As String
Dim yuklenen As Long
Dim FL As Integer
Dim A As String, C As String, D As String, E As String

Private Sub Command1_Click()
If Winsock1.State = 7 Then
Winsock1_Close
Else
Winsock1.Close
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = Text2.Text
Winsock1.Connect
End If
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
StatusBar1.Panels(3).ToolTipText = StatusBar1.Panels(3).Text
End Sub

Private Sub Winsock1_Close()
Command1.Caption = "Connect"
StatusBar1.Panels(1).Text = Winsock1.RemoteHost & " Disconnected.."
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
Command1.Caption = "Disconnect"
StatusBar1.Panels(1).Text = Winsock1.RemoteHost & " Connected.."
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

Dim thedata As String
Dim Ret As Integer

Winsock1.GetData thedata
If thedata = "" Then
Exit Sub

ElseIf Mid(thedata, 1, 6) = "<[70]>" Then ' "Msg_Eof_"
Close #FL 'Mid(thedata, 7, Len(thedata) - 3)

'Unload Form3
Form2.Show
Form2.Picture1.Picture = LoadPicture("")
Form2.Picture1.Refresh
Form2.Picture1.Picture = LoadPicture(Dosya)

FileBar.Refresh
yuklenen = 0

ElseIf Mid(thedata, 1, 6) = "<[71]>" Then ' "Msg_Dst_"

FileBar.Refresh
Dosya = Mid(thedata, 7, Len(thedata) - 3)
FL = FreeFile
On Error Resume Next
        If Len(Dir(Dosya)) > 0 Then
            Kill Dosya
        End If
Open Dosya For Binary As #FL
lbFilereceived.Caption = Dosya
Winsock1.SendData "<[80]>" ' "Msg_OkS"
DoEvents

ElseIf Mid(thedata, 1, 6) = "<[72]>" Then ' "Msg_Dst_"
FileBar.Refresh
FileBar.Min = 0
FileBar.Max = Str(Mid(thedata, 7, Len(thedata) - 3))
Winsock1.SendData "<[82]>"
DoEvents

Else
yuklenen = yuklenen + Len(thedata)
Put #FL, , thedata
lbBytesReceived.Caption = "Bytes received: " & yuklenen & " of " & FileBar.Max
FileBar.Value = yuklenen
StatusBar1.Panels(3).Text = "Received : % " & Int(yuklenen / FileBar.Max * 100)
Winsock1.SendData "<[81]>" ' "Msg_Rec"
DoEvents
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
StatusBar1.Panels(1).Text = Winsock1.RemoteHost & " Error : " & Description
Winsock1.Close
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Check1_Click()
If Check1.Value = 1 Then
Slider3.Enabled = True
Else
Slider3.Enabled = False
End If
End Sub

Private Sub Form_Load()
Show

Me.top = 0
Me.left = Screen.Width / 2 - Me.Width / 2

StatusBar1.Panels.Add
StatusBar1.Panels.Add
StatusBar1.Height = 350
StatusBar1.Panels(1).AutoSize = 1
StatusBar1.Panels(2).AutoSize = 1
StatusBar1.Panels(3).AutoSize = 1
Form2.Show

Slider1.Value = 2
Slider1.Text = "% " & Slider1.Value * 10
Slider2.Value = 4
Slider2.Text = Slider2.Value * 600
Slider3.Value = 4
Slider3.Text = Slider3.Value * 5

Form2.ScaleHeight = Screen.Height * Str(Slider1.Value) / 10
Form2.ScaleWidth = Screen.Width * Str(Slider1.Value) / 10
Slider3.Enabled = False

C = Label3 & Slider1.Text
D = "Mouse Move Length : " & Slider2.Text & " Bytes.."
E = Check1.Caption & Slider3.Text & " Seconds.."



KAYDIR
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
If Winsock1.State = 7 Then
    If Check1.Value = 1 Then
        Form1.Winsock1.SendData "<[51]>" & Slider1.Value & "^" & Slider3.Text
        DoEvents
    Else
        Form1.Winsock1.SendData "<[51]>" & Slider1.Value
    End If
Else
    StatusBar1.Panels(3).Text = "YOU ARE NOT CONNECT"
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next

SavePicture Form2.Picture1.Image, App.Path & "\" & Slider1.Text & ".bmp"
k = App.Path & "\" & Slider1.Text & ".bmp"
StatusBar1.Panels(3).Text = k & " File Saved.."
End Sub

Private Sub Slider1_Scroll()
Slider1.Text = "% " & Slider1.Value * 10
StatusBar1.Panels(3).Text = Label3 & Slider1.Text
C = Label3 & Slider1.Text
End Sub

Private Sub Slider2_Scroll()
Slider2.Text = Slider2.Value * 600
StatusBar1.Panels(3).Text = "Mouse Move Length : " & Slider2.Text & " Bytes.."
D = "Mouse Move Length : " & Slider2.Text & " Bytes.."
End Sub

Private Sub Slider3_Scroll()
Slider3.Text = Slider3.Value * 5
StatusBar1.Panels(3).Text = Check1.Caption & Slider3.Text & " Seconds.."
E = Check1.Caption & Slider3.Text & " Seconds.."
End Sub

Private Sub KAYDIR()
Dim B, i

Select Case A
    Case Is = C
        A = D
        B = Len(D)
    Case Is = D
        A = E
        B = Len(E)
    Case Is = E
        A = C
        B = Len(C)
    Case Else
        A = C
        B = Len(C)
End Select

    For i = 0 To B
        StatusBar1.Panels(2).Text = left(A, i)
        t = Timer
        Do: DoEvents: Loop Until Timer > t + 1 / 10
    Next
        Do: DoEvents: Loop Until Timer > t + 2
        
StatusBar1.Panels(3).Text = ""

KAYDIR
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   1560
   ClientLeft      =   8460
   ClientTop       =   6915
   ClientWidth     =   3420
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3420
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1305
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar FileBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbByteSend 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Byte Sended:"
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
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
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

Private Declare Function SetCursorPos& Lib "user32" _
(ByVal x As Long, ByVal y As Long)

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

Dim SrcPath As String
Dim DstPath As String
Dim IsReceived As Boolean

Dim f, h, v, b, mm, mw, mh, zz As String, kk As String, tt As Boolean
Private Sub SendFile()
    Dim BufFile As String
    Dim LnFile As Long
    Dim nLoop As Long
    Dim nRemain As Long
    Dim Cn As Long
    On Error GoTo GLocal:
    LnFile = FileLen(SrcPath)
        FileBar.Min = 0
        FileBar.Max = Str(LnFile)
            Winsock1.SendData "<[72]>" & Str(LnFile)
            IsReceived = False
            While IsReceived = False
                DoEvents
            Wend
    If LnFile > 8192 Then
        nLoop = Fix(LnFile / 8192)
        
        nRemain = LnFile Mod 8192
    Else
        nLoop = 0
        nRemain = LnFile
    End If
    
    If LnFile = 0 Then
        MsgBox "Ivalid Source File", vbCritical, "Client Message"
        Exit Sub
    End If
    
    Open SrcPath For Binary As #1
    If nLoop > 0 Then
        For Cn = 1 To nLoop
            BufFile = String(8192, " ")
            Get #1, , BufFile
            Winsock1.SendData BufFile
            IsReceived = False
            lbByteSend.Caption = "Bytes Sent: " & Cn * 8192 & " Of " & LnFile
            FileBar.Value = Str(Cn * 8192)
            StatusBar1.Panels(1).Text = "Screen Sent : % " & Int(FileBar.Value / FileBar.Max * 100)
            While IsReceived = False
                DoEvents
            Wend
        Next
        If nRemain > 0 Then
            BufFile = String(nRemain, " ")
            Get #1, , BufFile
            Winsock1.SendData BufFile
            IsReceived = False
            lbByteSend.Caption = "Bytes Sent: " & LnFile & " Of " & LnFile
            FileBar.Value = Str(LnFile)
            StatusBar1.Panels(1).Text = "Screen Sent : % " & Int(FileBar.Value / FileBar.Max * 100)
            While IsReceived = False
                DoEvents
            Wend
        End If
    Else
        BufFile = String(nRemain, " ")
        Get #1, , BufFile
        Winsock1.SendData BufFile
        IsReceived = False
        While IsReceived = False
            DoEvents
        Wend
    End If
    Winsock1.SendData "<[70]>"
    DoEvents

    Close #1
    Unload Form2
    FileBar.Refresh
    Exit Sub
GLocal:
    MsgBox Err.Description

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Form_Load()
If Winsock1.State = 2 Then
Winsock1.Close
Else
Winsock1.LocalPort = 1453
Winsock1.Listen
End If
SrcPath = App.Path & "\Screen.bmp"
DstPath = "c:\Screen.bmp"
h = True
tt = False

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 500

StatusBar1.Height = 350
StatusBar1.Panels(1).AutoSize = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Winsock1_Close()
StatusBar1.Panels(1).Text = Winsock1.RemoteHostIP & " Disconnected.."
Winsock1.Close
Form_Load
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
StatusBar1.Panels(1).Text = Winsock1.RemoteHostIP & " Connecting.."
Winsock1.Accept requestID
If Winsock1.State = 7 Then
StatusBar1.Panels(1).Text = Winsock1.RemoteHostIP & " Connected.."
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
On Error Resume Next

Dim thedata As String
Winsock1.GetData thedata

If Mid(thedata, 1, 5) = "<[1]>" Then
    MsgBox (Mid(thedata, 6, Len(thedata) - 3))

ElseIf Mid(thedata, 1, 6) = "<[51]>" Then
Dim jj As String
kk = (Mid(thedata, 7, Len(thedata) - 3))
jj = InStr(kk, "^")
tt = False
    If jj = 0 Then
        tt = False
    Else
        tt = True
        zz = Right(kk, Len(kk) - jj)
        kk = Left(kk, Len(kk) - Len(zz) - 1)
    End If
Call dfg

ElseIf Mid(thedata, 1, 6) = "<[80]>" Then
SendFile
ElseIf Mid(thedata, 1, 6) = "<[81]>" Then
IsReceived = True

ElseIf Mid(thedata, 1, 6) = "<[82]>" Then
IsReceived = True

ElseIf Mid(thedata, 1, 6) = "<[91]>" Then
Dim gb, xpos, ypos
gb = (Mid(thedata, 7, Len(thedata) - 3))
xpos = (Left(gb, InStr(gb, "|") - 1))
ypos = (Right(gb, Len(gb) - Len(xpos) - 1))

k = SetCursorPos&(xpos, ypos)
LeftClick
DoEvents

ElseIf Mid(thedata, 1, 6) = "<[92]>" Then
gb = (Mid(thedata, 7, Len(thedata) - 3))
xpos = (Left(gb, InStr(gb, "|") - 1))
ypos = (Right(gb, Len(gb) - Len(xpos) - 1))

k = SetCursorPos&(xpos, ypos)
RightClick
DoEvents

ElseIf Mid(thedata, 1, 6) = "<[93]>" Then
gb = (Mid(thedata, 7, Len(thedata) - 3))
xpos = (Left(gb, InStr(gb, "|") - 1))
ypos = (Right(gb, Len(gb) - Len(xpos) - 1))

k = SetCursorPos&(xpos, ypos)
LeftClick
DoEvents
LeftClick
DoEvents

ElseIf Mid(thedata, 1, 6) = "<[94]>" Then
gb = (Mid(thedata, 7, Len(thedata) - 3))
xpos = (Left(gb, InStr(gb, "|") - 1))
ypos = (Right(gb, Len(gb) - Len(xpos) - 1))

k = SetCursorPos&(xpos, ypos)
DoEvents

ElseIf Mid(thedata, 1, 6) = "<[95]>" Then
jjj = Mid(thedata, 7, Len(thedata) - 3)
mm = Right(jjj, Len(jjj) - InStr(jjj, " "))
mw = Left(mm, InStr(mm, "-") - 1)
mh = Right(mm, Len(mm) - InStr(mm, "-"))
jjj = Left(jjj, InStr(jjj, " ") - 1)
moving_ok jjj
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
StatusBar1.Panels(1).Text = ("Error : " & Description)
End Sub


Private Sub LeftClick()
    LeftDown
    LeftUp
End Sub

Private Sub RightClick()
    RightDown
    RightUp
End Sub

Private Sub LeftDown()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub LeftUp()
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub RightDown()
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
End Sub

Private Sub RightUp()
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Private Sub moving_ok(pos)
Dim a, s, d, g
s = Trim(pos)
d = InStr(s, "|")
If d = "0" Then Exit Sub
f = Left(s, d - 1)
GoToX
g = Right(s, Len(pos) - d)
pos = g
moving_ok (pos)
End Sub

Private Sub GoToX()
If h = False Then
GoToY
Exit Sub
Else
'MsgBox f, , " X"
Dim a, c
a = (Screen.Width / Screen.TwipsPerPixelX) / mw * f
c = (Screen.Height / Screen.TwipsPerPixelY) / mh * b
k = SetCursorPos&(a, c)
DoEvents
Pause (0.5)
v = f 'X
h = False
End If
End Sub

Private Sub GoToY()
'MsgBox f, , " Y"
Dim a, c
a = (Screen.Width / Screen.TwipsPerPixelX) / mw * v
c = (Screen.Height / Screen.TwipsPerPixelY) / mh * f

k = SetCursorPos&(a, c)
DoEvents
Pause (0.5)

b = f 'Y
h = True
End Sub

Private Sub Pause(time As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= time
        DoEvents
    Loop
End Sub

Private Sub dfg()
Form2.Picture1.Height = Screen.Height * kk / 10
Form2.Picture1.Width = Screen.Width * kk / 10
Form2.Show
Call Form2.yakala
Form2.Picture1.PaintPicture Form2.Picture, 0, 0, Form2.Picture1.Width, Form2.Picture1.Height, 0, 0, Screen.Width, Screen.Height, vbSrcCopy
DoEvents
SavePicture Form2.Picture1.Image, SrcPath
DoEvents
''''''''''''''''''''''''''''''''''''
On Error GoTo GLocal
Winsock1.SendData "<[71]>" & DstPath
DoEvents
    If tt = True Then
        Pause (Val(zz))
        DoEvents
        dfg
    Else
        tt = False
        Exit Sub
    End If
Exit Sub
GLocal:
StatusBar1.Panels(1).Text = ("Error : " & Description)
Resume Next
End Sub

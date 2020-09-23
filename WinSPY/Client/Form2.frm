VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Capture"
   ClientHeight    =   3570
   ClientLeft      =   15
   ClientTop       =   3615
   ClientWidth     =   5025
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   5025
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         Height          =   255
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Menu MnMouse 
      Caption         =   "Mouse"
      Begin VB.Menu MnStart 
         Caption         =   "Start"
      End
      Begin VB.Menu MnStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu MnKesme 
         Caption         =   "-"
      End
      Begin VB.Menu MnLClick 
         Caption         =   "Left Click"
      End
      Begin VB.Menu MnRClick 
         Caption         =   "Right Click"
      End
      Begin VB.Menu MnDblClick 
         Caption         =   "DoubleClick"
      End
      Begin VB.Menu MnMousePos 
         Caption         =   "Send Mouse Position"
      End
      Begin VB.Menu MnMove 
         Caption         =   "Move"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Osman Gulercan, Istanbul, TURKEY
'osmangulercan@yahoo.com
'o.gulercan@veezy.com

Private Type RECT
left As Integer
top As Integer
right As Integer
bottom As Integer
End Type
Private Type POINT
x As Long
y As Long
End Type

Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long)
Private Declare Function SetCursorPos& Lib "user32" (ByVal x As Long, ByVal y As Long)


Dim durum As Boolean
Dim moving As Boolean
Dim pos As String
Dim F, H, V, B

Private Sub Form_Load()
Show

Form2.left = Form1.left + Form1.Width / 2 - Form2.Width / 2
Form2.top = Form1.Height

H = True
moving = False
MnMouse.Visible = False
MnStop_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form2.left = Form1.left + Form1.Width / 2 - Form2.Width / 2
Form2.top = Form1.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
ClipCursor ByVal 0&
End Sub

Private Sub MnDblClick_Click()
moving = False
MnDblClick.Checked = True
MnLClick.Checked = False
MnMousePos.Checked = False
MnMove.Checked = False
MnRClick.Checked = False

Dim A, C
A = (Screen.Width / Screen.TwipsPerPixelY) / Picture1.ScaleWidth * (Shape1.left + Shape1.Width / 2)
C = (Screen.Height / Screen.TwipsPerPixelX) / Picture1.ScaleHeight * (Shape1.top + Shape1.Height / 2)
Form1.Winsock1.SendData "<[93]>" & (A) & "|" & (C)
DoEvents
End Sub

Private Sub MnLClick_Click()
moving = False
MnLClick.Checked = True
MnDblClick.Checked = False
MnMousePos.Checked = False
MnMove.Checked = False
MnRClick.Checked = False

Dim A, C
A = (Screen.Width / Screen.TwipsPerPixelY) / Picture1.ScaleWidth * (Shape1.left + Shape1.Width / 2)
C = (Screen.Height / Screen.TwipsPerPixelX) / Picture1.ScaleHeight * (Shape1.top + Shape1.Height / 2)
Form1.Winsock1.SendData "<[91]>" & (A) & "|" & (C)
DoEvents
End Sub

Private Sub MnMousePos_Click()
moving = False
MnMousePos.Checked = True
MnLClick.Checked = False
MnDblClick.Checked = False
MnMove.Checked = False
MnRClick.Checked = False

Dim A, C
A = (Screen.Width / Screen.TwipsPerPixelY) / Picture1.ScaleWidth * (Shape1.left + Shape1.Width / 2)
C = (Screen.Height / Screen.TwipsPerPixelX) / Picture1.ScaleHeight * (Shape1.top + Shape1.Height / 2)
Form1.Winsock1.SendData "<[94]>" & (A) & "|" & (C)
DoEvents
End Sub

Private Sub MnMove_Click()
moving = True
MnMove.Checked = True
MnLClick.Checked = False
MnDblClick.Checked = False
MnMousePos.Checked = False
MnRClick.Checked = False
'Your Mouse Position Recording & Sending in This Window;
'You can change mouse move slider;
'Please Move Your Mouse..
End Sub

Private Sub MnRClick_Click()
moving = False
MnRClick.Checked = True
MnDblClick.Checked = False
MnMousePos.Checked = False
MnMove.Checked = False
MnLClick.Checked = False

Dim A, C
A = (Screen.Width / Screen.TwipsPerPixelY) / Picture1.ScaleWidth * (Shape1.left + Shape1.Width / 2)
C = (Screen.Height / Screen.TwipsPerPixelX) / Picture1.ScaleHeight * (Shape1.top + Shape1.Height / 2)
Form1.Winsock1.SendData "<[92]>" & (A) & "|" & (C)
DoEvents
End Sub

Private Sub MnStart_Click()
moving = False
MnStart.Enabled = False
MnStop.Enabled = True
MnDblClick.Enabled = True
MnLClick.Enabled = True
MnRClick.Enabled = True
MnMousePos.Enabled = True
MnMove.Enabled = True

Dim client As RECT
Dim upperleft As POINT

GetClientRect Me.hWnd, client
upperleft.x = client.left
upperleft.y = client.top
ClientToScreen Me.hWnd, upperleft
OffsetRect client, upperleft.x, upperleft.y
ClipCursor client
End Sub

Private Sub MnStop_Click()
moving = False
MnStart.Enabled = True
MnStop.Enabled = False
MnDblClick.Enabled = False
MnLClick.Enabled = False
MnRClick.Enabled = False
MnMousePos.Enabled = False
MnMove.Enabled = False

MnMove.Checked = False
MnLClick.Checked = False
MnRClick.Checked = False
MnDblClick.Checked = False
MnMousePos.Checked = False

ClipCursor ByVal 0&
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
Shape1.left = x - Shape1.Width / 2
Shape1.top = y - Shape1.Height / 2
durum = False
End If
End Sub
Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
If durum = True Then
Shape1.left = x - Shape1.Width / 2
Shape1.top = y - Shape1.Height / 2
End If
End If

If moving = True Then
Shape1.left = x - Shape1.Width / 2
Shape1.top = y - Shape1.Height / 2
pos = pos & x & "|" & y & "|"
Caption = Len(pos)
If Len(pos) > Form1.Slider2.Text Then
'moving_ok
Form1.Winsock1.SendData "<[95]>" & pos & " " & Picture1.ScaleWidth & "-" & Picture1.ScaleHeight
DoEvents
moving = False
pos = ""
End If
End If
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
durum = True
End If

If Button = vbRightButton Then
PopupMenu MnMouse, 8, x, y
End If
End Sub
Private Sub Picture1_Resize()
On Error Resume Next
Form2.Width = Picture1.ScaleWidth + 100
Form2.Height = Picture1.ScaleHeight + 500
End Sub


VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2205
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   2205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   30
      ScaleHeight     =   1905
      ScaleWidth      =   2115
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   2145
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   1410
         Top             =   1095
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Reminder"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1545
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "13:15:00"
         Top             =   1095
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Remind At"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   855
         Width           =   1785
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   165
         TabIndex        =   9
         Text            =   "10"
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Remind In"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Width           =   1875
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   -15
         TabIndex        =   13
         Top             =   15
         Width           =   165
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Minutes"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1185
         TabIndex        =   10
         Top             =   555
         Width           =   555
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   345
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   930
      TabIndex        =   6
      ToolTipText     =   "Set Reminders"
      Top             =   75
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   3
      Left            =   1260
      TabIndex        =   5
      ToolTipText     =   "Hide Till Reminder"
      Top             =   60
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   1515
      TabIndex        =   4
      ToolTipText     =   "Set Window Position"
      Top             =   75
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   75
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   1695
      TabIndex        =   2
      ToolTipText     =   "Resize"
      Top             =   75
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1965
      TabIndex        =   0
      Top             =   1770
      Width           =   165
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTCAPTION = 2



Dim OldHeight As Long
Dim OldTime As Long
Private Sub Form_Load()
On Error Resume Next

Trans.MakeTransparent Me.hWnd, Form1.Slider1.Value

For Each ctl In Me
ctl.BackColor = Form1.Picture1.BackColor
ctl.ForeColor = Form1.Picture2.BackColor
Next ctl

Me.BackColor = Form1.Picture1.BackColor
Me.ForeColor = Form1.Picture2.BackColor

If Form1.Check1.Value = 1 Then
Label2(2).Caption = 2
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End If


With Label1
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
End With

Form_Resize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        'Release capture
        Call ReleaseCapture
        'Send a 'left mouse button down on caption'-message to our form
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = vbLeftButton Then
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
  End If

End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.MousePointer = 8
End Sub

Private Sub Form_Resize()
On Error Resume Next
Label1.Move Me.ScaleLeft + Me.ScaleWidth - (Label1.Width + 40), Me.ScaleTop + Me.ScaleHeight - (Label1.Height + 40)
Text1.Width = Me.Width - 500
Text1.Height = Me.Height - Label1.Height - 380

Label2(0).Left = Me.Width - 500
Label2(1).Left = Label2(0).Left - 200
Label2(2).Left = Label2(1).Left - 200
Label2(3).Left = Label2(2).Left - 200
Label2(4).Left = Label2(3).Left - 200
End Sub


Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(Index).FontBold = True

End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(Index).FontBold = False


Select Case Index

Case 0 ' x button
Unload Me

Case 1
If Me.Height = 360 Then
Me.Height = OldHeight
Else
OldHeight = Me.Height
Me.Height = 360
End If

Case 2
If Label2(2).Caption = 1 Then   '1 = ontop
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Label2(2).Caption = 2
Else
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Label2(2).Caption = 1
End If


Case 3
Me.Visible = False

Case 4
Picture1.Visible = True


End Select


End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = True
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = False
Picture1.Visible = False
End Sub

Private Sub Option1_Click()
Timer1.Interval = 60000
Timer1.Enabled = True
End Sub

Private Sub Option2_Click()
Timer1.Interval = 1000
Timer1.Enabled = True
End Sub

Private Sub Option3_Click()
Timer1.Enabled = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        'Release capture
        Call ReleaseCapture
        'Send a 'left mouse button down on caption'-message to our form
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub



Private Sub Timer1_Timer()

If Option2.Value = True Then
If Time = Text3.Text Then
Me.Visible = True
Picture1.Visible = False
Timer1.Enabled = False
Option3.Value = True
End If
End If

If Option1.Value = True Then
Text2 = Text2 - 1
If Text2.Text = 0 Then
Me.Visible = True
Picture1.Visible = False
Timer1.Enabled = False
Option3.Value = True
End If
End If




End Sub

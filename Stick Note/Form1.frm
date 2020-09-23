VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Stick Note ++"
   ClientHeight    =   1650
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Sticky.ShellIcon ShellIcon1 
      Left            =   2310
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      Icon            =   "Form1.frx":0442
      SysMenu         =   0   'False
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Stick Not Visibility"
      Top             =   1005
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   327682
      Max             =   255
      SelStart        =   255
      TickStyle       =   3
      TickFrequency   =   10
      Value           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   270
      Left            =   150
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set Stick Note On Top"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1605
      TabIndex        =   3
      Top             =   150
      Width           =   1950
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   765
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   2
      ToolTipText     =   "Text Colour"
      Top             =   390
      Width           =   390
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   30
      Top             =   2475
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   225
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      ToolTipText     =   "Back Colour"
      Top             =   390
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Opaque"
      Height          =   195
      Left            =   3135
      TabIndex        =   7
      Top             =   1230
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Invisible"
      Height          =   195
      Left            =   1620
      TabIndex        =   6
      Top             =   1230
      Width           =   570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sticky Note Colour"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   1575
   End
   Begin VB.Menu main 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu backcol 
         Caption         =   "Back Colour"
      End
      Begin VB.Menu txtcol 
         Caption         =   "Text Colour"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu loadvis 
         Caption         =   "Load"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu restoring 
         Caption         =   "Restore"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backcol_Click()
Picture1_Click
End Sub

Private Sub Command1_Click()
Dim Gb As Form2
Set Gb = New Form2
Load Gb
Gb.Visible = True
End Sub

Private Sub exit_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
SendBottomRight Form1
ShellIcon1.ToolTipText = "Sticky Note ++"
ShellIcon1.Visible = True

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
Me.Hide
Else
Me.Show
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShellIcon1.Visible = False
Unload Me
End
End Sub

Private Sub loadvis_Click()
Command1_Click
End Sub

Private Sub Picture1_Click()
On Error Resume Next
cd1.ShowColor
Picture1.BackColor = cd1.Color
End Sub
Private Sub Picture2_Click()
On Error Resume Next
cd1.ShowColor
Picture2.BackColor = cd1.Color

End Sub

Private Sub restoring_Click()
Me.WindowState = 0
Me.Show
End Sub

Private Sub ShellIcon1_MouseUp(Button As Integer)
If Button = vbRightButton Then PopupMenu main
End Sub

Private Sub txtcol_Click()
Picture2_Click
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WindowsXP SideMenu"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   3240
      Top             =   1440
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      LargeChange     =   500
      Left            =   3120
      Max             =   1000
      SmallChange     =   500
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000D&
         Height          =   3615
         Left            =   0
         Picture         =   "Form1.frx":1042
         ScaleHeight     =   3615
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   0
         Width           =   3135
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   50
            Left            =   120
            Picture         =   "Form1.frx":49DA4
            ScaleHeight     =   45
            ScaleWidth      =   2775
            TabIndex        =   11
            Top             =   600
            Width           =   2775
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Control Panel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "Form1.frx":6823A
               MousePointer    =   99  'Custom
               TabIndex        =   14
               ToolTipText     =   "Change the computer settings"
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Set Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "Form1.frx":69FBC
               MousePointer    =   99  'Custom
               TabIndex        =   13
               ToolTipText     =   "Set the time"
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Taskbar Settings"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   480
               MouseIcon       =   "Form1.frx":6BD3E
               MousePointer    =   99  'Custom
               TabIndex        =   12
               ToolTipText     =   "Taskbar Settigns"
               Top             =   600
               Width           =   1335
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   120
               Picture         =   "Form1.frx":6DAC0
               Top             =   120
               Width           =   240
            End
            Begin VB.Image Image2 
               Height          =   240
               Left            =   120
               Picture         =   "Form1.frx":6DE02
               Top             =   600
               Width           =   240
            End
            Begin VB.Image Image3 
               Height          =   270
               Left            =   120
               Picture         =   "Form1.frx":6E144
               Top             =   360
               Width           =   285
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2520
            MouseIcon       =   "Form1.frx":6E5BE
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":70340
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   10
            Top             =   240
            Width           =   375
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   400
            Left            =   120
            ScaleHeight     =   405
            ScaleWidth      =   2775
            TabIndex        =   2
            Top             =   800
            Width           =   2775
            Begin VB.PictureBox Picture8 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   0
               MouseIcon       =   "Form1.frx":70AEE
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":72870
               ScaleHeight     =   375
               ScaleWidth      =   2775
               TabIndex        =   8
               Top             =   0
               Width           =   2775
               Begin VB.Label Label8 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Find"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   240
                  MouseIcon       =   "Form1.frx":75EFE
                  MousePointer    =   99  'Custom
                  TabIndex        =   9
                  Top             =   120
                  Width           =   1455
               End
            End
            Begin VB.PictureBox Picture10 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2400
               MouseIcon       =   "Form1.frx":77C80
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":79A02
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   7
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   0
               Picture         =   "Form1.frx":7A1B0
               ScaleHeight     =   735
               ScaleWidth      =   2775
               TabIndex        =   4
               Top             =   360
               Width           =   2775
               Begin VB.Image Image4 
                  Height          =   240
                  Left            =   120
                  Picture         =   "Form1.frx":98646
                  Top             =   360
                  Width           =   225
               End
               Begin VB.Image Image6 
                  Appearance      =   0  'Flat
                  Height          =   240
                  Left            =   120
                  Picture         =   "Form1.frx":98988
                  Top             =   120
                  Width           =   225
               End
               Begin VB.Label Label6 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Computers"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   480
                  MouseIcon       =   "Form1.frx":98CCA
                  MousePointer    =   99  'Custom
                  TabIndex        =   6
                  ToolTipText     =   "Find Computers"
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label Label7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Files"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   480
                  MouseIcon       =   "Form1.frx":9AA4C
                  MousePointer    =   99  'Custom
                  TabIndex        =   5
                  ToolTipText     =   "Find the files on Your computer"
                  Top             =   120
                  Width           =   375
               End
            End
            Begin VB.PictureBox Picture9 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   2400
               MouseIcon       =   "Form1.frx":9C7CE
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":9E550
               ScaleHeight     =   375
               ScaleWidth      =   375
               TabIndex        =   3
               Top             =   0
               Width           =   375
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2520
            MouseIcon       =   "Form1.frx":9ECFE
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":A0A80
            ScaleHeight     =   375
            ScaleWidth      =   375
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            MouseIcon       =   "Form1.frx":A122E
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":A2FB0
            ScaleHeight     =   375
            ScaleWidth      =   2775
            TabIndex        =   16
            Top             =   240
            Width           =   2775
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "System Tasks"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   240
               MouseIcon       =   "Form1.frx":A663E
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   120
               Width           =   1455
            End
         End
      End
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   960
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   960
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   3960
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sh As New Shell 'This is to acces shell automation
Private Sub moff()
Timer2.Enabled = True ' This disables the menu animation
End Sub

Private Sub mon()
Timer1.Enabled = True ' This enables the menu animation
End Sub

Private Sub mnon()
Timer6.Enabled = True ' This enables the 2nd menu animation
End Sub

Private Sub mnoff()
Timer7.Enabled = True ' This disables the 2nd menu animation
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label1_Click()
sh.ControlPanelItem (a) 'Uses Shell automation to show Control Panel
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This controls Behaviour of some label controls
Label1.FontUnderline = True 'underlines label1
Label1.ForeColor = &HFF8080 'changes colour of label 1
Label3.FontUnderline = False 'removes underline of label3
Label4.FontUnderline = False 'removes underline of label3
Label3.ForeColor = &HFF0000 '1
Label4.ForeColor = &HFF0000 '2
'These two line change colour of other labels on the menu picture box
End Sub
Private Sub Label2_Click()
If Picture5.Visible = True Then
Call mon
'shows menu if down arrow (picture5) is clicked
Picture5.Visible = False 'hides down arrow
Picture4.Visible = True 'Shows up arrow

ElseIf Picture4.Visible = True Then
Call moff
'hides menu if up arrow (picture4) is click
Picture4.Visible = False 'hides up arrow
Picture5.Visible = True 'Shows down arrow
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF8080 'if mouse is over label2 it changes its colour
End Sub

Private Sub Label3_Click()
sh.SetTime 'Uses shell automation to show system time settings.
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontUnderline = True 'underlines label3
Label3.ForeColor = &HFF8080 'changes fore colour of label1
Label1.FontUnderline = False 'removes underline
Label4.FontUnderline = False 'removes underline
Label1.ForeColor = &HFF0000 'changes forecolor
Label4.ForeColor = &HFF0000 ' chnges forecolor
End Sub

Private Sub Label4_Click()
sh.TrayProperties ' Use shell automation to show taskbar properties
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontUnderline = True 'underlines this label
Label4.ForeColor = &HFF8080 'changes forecolor
Label1.FontUnderline = False 'removes underline
Label3.FontUnderline = False 'removes underline
Label1.ForeColor = &HFF0000 'changes forecolor
Label3.ForeColor = &HFF0000 'changes forecolor
End Sub

Private Sub Label6_Click()
sh.FindComputer 'Find a computer Dialog
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontUnderline = True 'underlines label3
Label6.ForeColor = &HFF8080 'changes fore colour of label1
Label7.FontUnderline = False 'removes underline
Label7.ForeColor = &HFF0000 ' chnges forecolor
End Sub

Private Sub Label7_Click()
sh.FindFiles 'Shows windows file find system

End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.FontUnderline = True 'underlines label3
Label7.ForeColor = &HFF8080 'changes fore colour of label1
Label6.FontUnderline = False 'removes underline
Label6.ForeColor = &HFF0000 ' changes forecolor
End Sub

Private Sub Label8_Click()
If Picture9.Visible = True Then
Call mnon
'shows menu if down arrow (picture5) is clicked
Picture9.Visible = False 'hides down arrow
Picture10.Visible = True 'Shows up arrow

ElseIf Picture10.Visible = True Then
Call mnoff
'hides menu if up arrow (picture4) is click
Picture10.Visible = False 'hides up arrow
Picture9.Visible = True 'Shows down arrow
End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF8080 'if mouse is over label8 it changes its colour
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This changes color when mouse is over picture1
Label2.ForeColor = &HFF0000 'changes forecolor
End Sub

Private Sub Picture10_Click()
Call mnoff
'hides 2nd menu
Picture10.Visible = False 'hides up arrow
Picture9.Visible = True 'shows down arrow
End Sub

Private Sub Picture11_Click()
If Picture5.Visible = True Then
Call mon
'shows menu
Picture5.Visible = False 'hides down arrow
Picture4.Visible = True 'shows up arrow
ElseIf Picture4.Visible = True Then
Call moff
'hides menu
Picture4.Visible = False 'hides up arrow
Picture5.Visible = True 'shows down arrow
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF0000 'changes forecolor

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF0000 'changes forecolor
Label1.ForeColor = &HFF0000 'changes forecolor
Label3.ForeColor = &HFF0000 'changes forecolor
Label4.ForeColor = &HFF0000 'changes forecolor
Label1.FontUnderline = False 'removes underline
Label3.FontUnderline = False 'removes underline
Label4.FontUnderline = False 'removes underline
End Sub

Private Sub Picture4_Click()
Call moff
'hides menu
Picture4.Visible = False 'hides up arrow
Picture5.Visible = True 'shows down arrow
End Sub

Private Sub Picture5_Click()
Call mon
'shows menu
Picture5.Visible = False 'hides down arrow
Picture4.Visible = True 'shows up arrow
End Sub


Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF0000 'changes forecolor
Label6.ForeColor = &HFF0000 'changes forecolor
Label7.ForeColor = &HFF0000 'changes forecolor
Label6.FontUnderline = False 'removes underline
Label7.FontUnderline = False 'removes underline
End Sub

Private Sub Picture8_Click()
If Picture9.Visible = True Then
Call mnon
'shows menu
Picture9.Visible = False 'hides down arrow
Picture10.Visible = True 'shows up arrow
ElseIf Picture10.Visible = True Then
Call mnoff
'hides menu
Picture10.Visible = False 'hides up arrow
Picture9.Visible = True 'shows down arrow
End If
End Sub

Private Sub Picture8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF0000 'changes forecolor
End Sub

Private Sub Picture9_Click()
Call mnon
'shows menu
Picture9.Visible = False 'hides down arrow
Picture10.Visible = True 'shows up arrow
End Sub

Private Sub Timer1_Timer()
'This turns on the menu animation
If Picture3.Height = 1000 Then
Timer1.Enabled = False
Else
Picture3.Height = Picture3.Height + 50
End If
'if height of 1st menu is 1000 turn off the timer else
'increase the height by 50
End Sub

Private Sub Timer2_Timer()
'this turns off the menu animation
If Picture3.Height = 50 Then
Timer2.Enabled = False
Else
Picture3.Height = Picture3.Height - 50
End If
'if height of 1st menu is 50 turn off the timer else
'decrease the height by 50
End Sub
'These menu animation speeds work well on windows XP but
'for Windows 98 they must be adjusted to work correctly.

Private Sub Timer3_Timer()
'this controls the movement of 2nd menu according to the first
'moves to downward
If Picture6.Top = 1800 Then
Timer3.Enabled = False
Else
Picture6.Top = Picture6.Top + 50
End If
End Sub


Private Sub Timer4_Timer()
'This move 2nd menu downwards
If Picture6.Top = 800 Then
Timer4.Enabled = False
Else
Picture6.Top = Picture6.Top - 50
End If
End Sub


Private Sub Timer5_Timer()
'Checks if 1st menu is contracted or expanded
'and makes the 2nd menu work accordingly
If Picture3.Height > 50 Then
Timer3.Enabled = True
ElseIf Picture3.Height < 1000 Then
Timer4.Enabled = True
End If
End Sub

Private Sub Timer6_Timer()
'This turns on the menu animation
If Picture6.Height = 1100 Then
Timer6.Enabled = False
Else
Picture6.Height = Picture6.Height + 50
End If
'if height of 2nd menu is 1100 turn off the timer else
'increase the height by 50
End Sub

Private Sub Timer7_Timer()
'this turns off the menu animation
If Picture6.Height = 400 Then
Timer7.Enabled = False
Else
Picture6.Height = Picture6.Height - 50
End If
'if height of 2nd is 400 turn off the timer else
'decrease the height by 50
End Sub

Private Sub Timer8_Timer()
'When both menus are expanded show the scrollbar
If Picture6.Height > 400 And Picture3.Height > 50 Then
VScroll1.Visible = True
Else
VScroll1.Visible = False
End If

End Sub

Private Sub VScroll1_Change()
Picture2.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_scroll()
Picture2.Top = -VScroll1.Value
End Sub

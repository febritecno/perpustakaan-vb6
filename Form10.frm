VERSION 5.00
Begin VB.Form Form10 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form10.frx":058A
   ScaleHeight     =   7920
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   840
      Top             =   8040
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   0
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   645
      IMEMode         =   3  'DISABLE
      Left            =   4440
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "LOG IN"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Shruti"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4440
      MaskColor       =   &H00C0FFFF&
      Picture         =   "Form10.frx":5D172
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   3855
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If UCase(Text1) = "ADMIN" And UCase(Text2) = "ADMIN" Then

    Command1.SetFocus
  Form9.Show
  retval = PlayWaveFile(App.Path & "\dooropen.wav", True)
  Form10.Hide
  Text1 = ""
  Text2 = ""
  Else
 MsgBox "Mohon Masukan Yang Benar ", vbOKOnly + vbInformation, "Notification"
 Text1 = ""
 Text2 = ""
 Text1.Enabled = True
 Text1.SetFocus
 End If
End Sub



Private Sub Form_Unload(cancel As Integer)
If MsgBox(" MAU KELUAR....??", vbQuestion + vbYesNo, "Info") = vbYes Then
retval = PlayWaveFile(App.Path & "\ROBOT.wav", True)
Animation

Form9.Show
Me.Hide

End
Else
cancel = 1
End If
End Sub
Public Sub Animation()
Dim I As Long
Dim J As Long
J = Me.ScaleHeight
I = Me.ScaleWidth


While Not I = 0
Me.Height = Me.Height - 25
I = I - 10
Wend


While Not J = 0
Me.Width = Me.Width - 25
J = J - 10
Wend
End Sub

    Private Sub Form_Load()
    Me.Height = 0
    retval = PlayWaveFile(App.Path & "\robot_explode.wav", True)
    End Sub

Private Sub Text1_Change()
If UCase(Text1) = "ADMIN" Then
Text2.SetFocus
Text1.Enabled = False
End If
End Sub


    Private Sub Timer1_Timer()
    Me.Height = Me.Height + 100
    Tengah
    If Me.Height >= 8355 Then
    Timer1.Enabled = False
    Tengah
    End If
    End Sub

    Public Sub Tengah()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    End Sub




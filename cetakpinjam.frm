VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CETAK PEMINJAMAN BUKU"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "CETAK LAPORAN PEMINJAMAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      MaskColor       =   &H00404040&
      Picture         =   "cetakpinjam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport5.Show
End Sub

Private Sub Form_Unload(cancel As Integer)
Do
    Me.Top = Me.Top + 40
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Top > Screen.Height - 500
    retval = PlayWaveFile(App.Path & "\ROBOT.wav", True)
Form9.Show
Me.Hide

End Sub

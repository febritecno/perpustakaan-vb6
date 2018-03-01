VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "APPS LIB"
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12735
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form9.frx":15A40
   ScaleHeight     =   7905
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   3360
      Top             =   7920
   End
   Begin VB.Menu File 
      Caption         =   "Fil&e"
      Begin VB.Menu A 
         Caption         =   "Pendataan Siswa"
      End
      Begin VB.Menu B 
         Caption         =   "Pendataan Kategori Buku"
      End
      Begin VB.Menu C 
         Caption         =   "Pendataan Petugas"
      End
      Begin VB.Menu D 
         Caption         =   "Pendataan Buku"
      End
      Begin VB.Menu MK 
         Caption         =   "-"
      End
      Begin VB.Menu E 
         Caption         =   "Peminjaman Buku"
      End
      Begin VB.Menu F 
         Caption         =   "Kartu Anggota"
      End
      Begin VB.Menu ZAB 
         Caption         =   "-"
      End
      Begin VB.Menu G 
         Caption         =   "Laporan Petugas"
      End
      Begin VB.Menu H 
         Caption         =   "Laporan Buku"
      End
      Begin VB.Menu I 
         Caption         =   "Laporan Peminjaman Buku"
      End
      Begin VB.Menu KB 
         Caption         =   "Laporan Kategori Buku"
      End
      Begin VB.Menu J 
         Caption         =   "-"
      End
      Begin VB.Menu K 
         Caption         =   "LOGOUT"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Z 
      Caption         =   "Hel&p"
      Begin VB.Menu ZX 
         Caption         =   "ME"
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub A_Click()
Form1.Show
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Hide
End Sub

Private Sub B_Click()
Form1.Hide
Form2.Show
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Hide
End Sub

Private Sub C_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Show
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Hide
End Sub

 Private Sub Form_Load()
    Me.Width = 0
    End Sub

    Private Sub Timer1_Timer()
    Me.Width = Me.Width + 100
    Tengah
    If Me.Width >= 12975 Then
    Timer1.Enabled = False
    Tengah
    End If
    End Sub

    Public Sub Tengah()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    End Sub


Private Sub D_Click()
Form1.Hide
Form2.Hide
Form3.Show
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Hide
End Sub

Private Sub E_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Show
End Sub

Private Sub F_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Show
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide
Form11.Hide
End Sub

Private Sub G_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form8.Show
Form11.Hide
End Sub

Private Sub H_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form7.Show
Form11.Hide
End Sub

Private Sub I_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form11.Hide
Form12.Show
End Sub

Private Sub K_Click()
Do
    Me.Top = Me.Top + 40
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Top > Screen.Height - 500
retval = PlayWaveFile(App.Path & "\ROBOT.wav", True)

Form10.Show
Form10.Text1.Enabled = True
Form10.Text1.SetFocus
End Sub

Private Sub KB_Click()
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form6.Show
Form11.Hide
End Sub

Private Sub Form_Unload(cancel As Integer)
Do
    Me.Top = Me.Top + 40
    Me.Move Me.Left, Me.Top
    DoEvents
    Loop Until Me.Top > Screen.Height - 500

retval = PlayWaveFile(App.Path & "\ROBOT.wav", True)
Me.Hide
Form10.Text1.Enabled = True
Form10.Show

Form10.Text1.SetFocus
End Sub

Private Sub ZX_Click()
Form13.Show
Form9.Hide
End Sub

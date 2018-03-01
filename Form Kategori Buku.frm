VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Kategori Buku"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12750
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form Kategori Buku.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Connect         =   "Access"
      DatabaseName    =   "G:\FEBRIAN\Perpustakaan\perpustakaan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Width           =   7095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      Picture         =   "Form Kategori Buku.frx":50A21
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6360
      Picture         =   "Form Kategori Buku.frx":51423
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      ItemData        =   "Form Kategori Buku.frx":51E25
      Left            =   6240
      List            =   "Form Kategori Buku.frx":51E44
      TabIndex        =   5
      Text            =   "Pilih Kategori"
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2880
      Picture         =   "Form Kategori Buku.frx":51EA5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT KATEGORI BUKU"
      BeginProperty Font 
         Name            =   "Danube"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   705
      Left            =   3045
      TabIndex        =   9
      Top             =   1200
      Width           =   6585
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Kategori : "
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   3000
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Kategori     : "
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   2565
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rskategori As Recordset
Public X
Dim pilih


Private Sub Command1_Click()
X = 1
papaw

End Sub

Private Sub Command2_Click()
Save
End Sub

Private Sub Command3_Click()
X = 2
Text2.Enabled = False
Combo1.SetFocus
End Sub

Private Sub Command4_Click()
bersih
End Sub

Private Sub Command5_Click()
pesan = MsgBox("Yakin Mau dihapus", vbYesNo + vbExclamation, "Warning")
If pesan = vbYes Then
rskategori.Delete
bersih
End If

End Sub


Private Sub Command6_Click()
Form2.Hide
Form9.Show
End Sub


Private Sub Form_Activate()
Text2.Enabled = False
Combo1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False

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

Private Sub Form_Load()
Set dbpustaka = OpenDatabase(App.Path & "\perpustakaan.mdb")
Set rskategori = dbpustaka.OpenRecordset("Kategori")
End Sub
Function papaw()
Text2.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Text2.SetFocus

End Function
Function bersih()
Text2.Text = ""
Combo1.Text = "[PILIH KATEGORI]"
'Text2.SetFocus
End Function
Function Save()
If X = 1 Then
rskategori.AddNew
rskategori!No_Kategori = Text2.Text
rskategori!Nama_Kategori = Combo1.Text
rskategori.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
ElseIf X = 2 Then
rskategori.Edit
rskategori!No_Kategori = Text2
rskategori!Nama_Kategori = Combo1.Text
rskategori.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
End If
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Function tampil()
Text2.Text = rskategori!No_Kategori
Combo1.Text = rskategori!Nama_Kategori

End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rskategori.Index = "No_Kategori"
rskategori.Seek "=", Text2.Text
If rskategori.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Combo2.SetFocus
Else
tampil
End If
End If
End Sub

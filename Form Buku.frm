VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form Buku"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12765
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form Buku.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   12765
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H8000000D&
      Height          =   390
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7545
      Width           =   7935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   7680
      TabIndex        =   20
      Top             =   6360
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
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
      Height          =   975
      Left            =   3120
      Picture         =   "Form Buku.frx":50A21
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000D&
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   1680
      Picture         =   "Form Buku.frx":51423
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
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
      Height          =   615
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
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
      Height          =   975
      Left            =   4440
      Picture         =   "Form Buku.frx":51E25
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   1335
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
      Height          =   975
      Left            =   1680
      Picture         =   "Form Buku.frx":52827
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   7680
      TabIndex        =   13
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   7680
      TabIndex        =   12
      Top             =   4920
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   7680
      TabIndex        =   11
      Top             =   4200
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   3600
      TabIndex        =   8
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   3600
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Input Data Buku"
      BeginProperty Font 
         Name            =   "Dabble(eval)"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   4605
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   6240
      TabIndex        =   21
      Top             =   6480
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   6240
      TabIndex        =   6
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pengarang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penerbit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   6240
      TabIndex        =   4
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Judul Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Buku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   1425
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rskategori As Recordset
Dim rsbuku As Recordset
Public X
Dim pilih


Private Sub Command1_Click()
X = 1
papaw

End Sub

Private Sub Command2_Click()
X = 2
Text1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Command3_Click()
bersih
End Sub

Private Sub Command4_Click()
pesan = MsgBox("Yakin Mau dihapus", vbYesNo + vbExclamation, "Warning")
If pesan = vbYes Then
rskategori.Delete
bersih
End If
End Sub

Private Sub Command5_Click()
Form3.Hide
Form9.Show
End Sub


Private Sub Command7_Click()
Save
End Sub

Private Sub Form_Activate()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = True
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
Set rsbuku = dbpustaka.OpenRecordset("buku")
End Sub
Function papaw()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
Text2.SetFocus

End Function
Function bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Function
Function Save()
If X = 1 Then
rsbuku.AddNew
rsbuku!No_Kategori = Text1.Text
rsbuku!kode_buku = Text3.Text
rsbuku!judul = Text4.Text
rsbuku!penerbit = Text5.Text
rsbuku!pengarang = Text6.Text
rsbuku!jumlah = Text7.Text
rsbuku!harga = Text8.Text
rsbuku.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
ElseIf X = 2 Then
rsbuku.Edit
rsbuku!No_Kategori = Text1.Text
rsbuku!kode_buku = Text3.Text
rsbuku!judul = Text4.Text
rsbuku!penerbit = Text5.Text
rsbuku!pengarang = Text6.Text
rsbuku!jumlah = Text7.Text
rsbuku!harga = Text8.Text
rsbuku.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
End If
End Function



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rskategori.Index = "No_Kategori"
rskategori.Seek "=", Text1.Text
If rskategori.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text3.SetFocus
Else
tampil
End If
End If
End Sub

Function tampil()
Text2.Text = rskategori!Nama_Kategori

End Function



Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rsbuku.Index = "kode_buku"
rsbuku.Seek "=", Text3.Text
If rsbuku.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text5.SetFocus
Else
tampil1
End If
End If
End Sub
Function tampil1()
Text1.Text = rsbuku!No_Kategori
Text2.Text = rskategori!Nama_Kategori
Text3.Text = rsbuku!kode_buku
Text4.Text = rsbuku!judul
Text5.Text = rsbuku!penerbit
Text6.Text = rsbuku!pengarang
Text7.Text = rsbuku!jumlah
Text8.Text = rsbuku!harga

End Function


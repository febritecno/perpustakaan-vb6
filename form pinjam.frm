VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PEMINJAMAN BUKU"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form pinjam.frx":0000
   ScaleHeight     =   7725
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Connect         =   "Access"
      DatabaseName    =   "G:\FEBRIAN\Perpustakaan\perpustakaan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Width           =   8775
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      Picture         =   "form pinjam.frx":715CD
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Picture         =   "form pinjam.frx":71FCF
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   8760
      TabIndex        =   27
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   8760
      TabIndex        =   26
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Ubah"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "AR JULIAN"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Picture         =   "form pinjam.frx":729D1
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   8760
      TabIndex        =   23
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   8760
      TabIndex        =   21
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Stok"
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
      Height          =   285
      Left            =   6000
      TabIndex        =   20
      Top             =   3840
      Width           =   1830
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Pinjam"
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
      Height          =   285
      Left            =   6000
      TabIndex        =   19
      Top             =   3120
      Width           =   2010
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Kembali"
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
      Height          =   285
      Left            =   6000
      TabIndex        =   18
      Top             =   2400
      Width           =   2505
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Pinjam"
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
      Height          =   285
      Left            =   6000
      TabIndex        =   17
      Top             =   1680
      Width           =   2265
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pinjam"
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
      Height          =   285
      Left            =   6000
      TabIndex        =   16
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Buku"
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
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   5760
      Width           =   1860
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Penerbit"
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
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Judul Buku"
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
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   4320
      Width           =   1680
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Buku"
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
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   3600
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Jurusan"
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
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
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
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Siswa"
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
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "NISN"
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
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BorderColor     =   &H000080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   5895
      Left            =   240
      Top             =   480
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BorderColor     =   &H000080FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   4095
      Left            =   5880
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rssiswa As Recordset
Dim rspinjam As Recordset
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
Text5.SetFocus
End Sub

Private Sub Command3_Click()
Save
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
Form11.Hide
Form9.Show
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




Private Sub Form_Activate()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End Sub

Private Sub Form_Load()
Set dbpustaka = OpenDatabase(App.Path & "\perpustakaan.mdb")
Set rssiswa = dbpustaka.OpenRecordset("Siswa")
Set rsbuku = dbpustaka.OpenRecordset("buku")
Set rspinjam = dbpustaka.OpenRecordset("pinjam")
End Sub
Function papaw()
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = True
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Text5.SetFocus

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
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End Function
Function Save()
If X = 1 Then
rspinjam.AddNew
rspinjam!nis = Text1.Text
rspinjam!kode_buku = Text5.Text
rspinjam!kode_pinjam = Text9.Text
rspinjam!tanggal_pinjam = Text10.Text
rspinjam!tanggal_kembali = Text11.Text
rspinjam!jumlah = Text12.Text
rspinjam!keterangan = Text13.Text
rspinjam.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
ElseIf X = 2 Then
rspinjam.Edit
rspinjam!nis = Text1.Text
rspinjam!kode_buku = Text5.Text
rspinjam!kode_pinjam = Text9.Text
rspinjam!tanggal_pinjam = Text10.Text
rspinjam!tanggal_kembali = Text11.Text
rspinjam!jumlah = Text12.Text
rspinjam!keterangan = Text13.Text
rsbuku.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
bersih
End If
End Function



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rssiswa.Index = "nis"
rssiswa.Seek "=", Text1.Text
If rssiswa.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text5.SetFocus
Else
tampil
End If
End If
End Sub

Function tampil()
Text2.Text = rssiswa!Nama
Text3.Text = rssiswa!kelas
Text4.Text = rssiswa!jurusan

End Function


Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text13.Text = Val(Text8.Text) - Val(Text12.Text)
If Text12 > Text8 Then
pesan = MsgBox("Maaf Jumlah yang anda masukkan terlalu besar", vbOKOnly, "Warning")
End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rsbuku.Index = "kode_buku"
rsbuku.Seek "=", Text5.Text
If rsbuku.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text5.SetFocus
Else
tampil1
End If
End If
End Sub
Function tampil1()
Text6.Text = rsbuku!judul
Text7.Text = rsbuku!penerbit
Text8.Text = rsbuku!jumlah

End Function


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rsbuku.Index = "kode_buku"
rsbuku.Seek "=", Text5.Text
If rsbuku.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text9.SetFocus
Else
tampil1
End If
End If
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rspinjam.Index = "kode_pinjam"
rspinjam.Seek "=", Text1.Text
If rspinjam.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text10.SetFocus
Else
tampil2
End If
End If
End Sub
Function tampil2()
Text10.Text = rspinjam!tanggal_pinjam
Text11.Text = rspinjam!tanggal_kembali
Text12.Text = rspinjam!jumlah
Text13.Text = rspinjam!keterangan

End Function

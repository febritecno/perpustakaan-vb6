VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Kategori"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "Kategori Buku.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Perpustakaan\Perpustakaan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ubah"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Kategori Buku.frx":038A
      Left            =   5520
      List            =   "Kategori Buku.frx":03A9
      TabIndex        =   3
      Text            =   "(pilih kategori)"
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nomer Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   1980
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rskategori As Recordset
Public x
Dim Pilih

Private Sub Command1_Click()
x = 1
ardin
End Sub

Private Sub Command2_Click()
Simpan
End Sub

Private Sub Command3_Click()
x = 2
Text1.Enabled = False
Combo1.SetFocus
End Sub

Private Sub Command4_Click()
Bersih
End Sub

Private Sub Command5_Click()
pesan = MsgBox("Yakin data akan dihapus?", vbYesNo + vbExclamation, "Warning")
If pesan = vbYes Then
    rskategori.Delete
    Bersih
End If
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Activate()
Text1.Enabled = False
Combo1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = True
End Sub

Private Sub Form_Load()
Set dbpustaka = OpenDatabase("D:\Perpustakaan\Perpustakaan.mdb")
Set rskategori = dbpustaka.OpenRecordset("kategori")
End Sub
Function ardin()
Text1.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Text1.SetFocus
End Function
Function Bersih()
Text1.Text = ""
Combo1.Text = "[Pilih Kategori]"
'Text1.SetFocus

End Function
Function Simpan()
If x = 1 Then
rskategori.AddNew
rskategori!nomor_kategori = Text1
rskategori!nama_kategori = Combo1.Text
rskategori.Update
pesan = MsgBox("Data Telah di Simpan", vbOKOnly + vbInformation, "Information")
Bersih
ElseIf x = 2 Then
rskategori.Edit
rskategori!nomor_kategori = Text1
rskategori!nama_kategori = Combo1.Text
rskategori.Update
pesan = MsgBox("Data Telah di Ubah", vbOKOnly + vbInformation, "Information")
Bersih
End If
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rskategori.Index = "nomor_kategori"
rskategori.Seek "=", Text1.Text
If rskategori.NoMatch Then
pesan = MsgBox(" Data Tidak Ada", vbOKOnly + vbInformation, "warning")
Combo1.SetFocus
Else
tampil
End If
End If


End Sub

Function tampil()
Text1.Text = rskategori!nomor_kategori
Combo1.Text = rskategori!nama_kategori

End Function


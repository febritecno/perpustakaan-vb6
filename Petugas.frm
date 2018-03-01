VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Petugas"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "Petugas.frx":0000
   LinkTopic       =   "Form4"
   Picture         =   "Petugas.frx":038A
   ScaleHeight     =   6225
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Petugas.frx":3AB3A0
      Left            =   2160
      List            =   "Petugas.frx":3AB3AA
      TabIndex        =   16
      Text            =   "(pilih jenis kelamin)"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Perpustakaan\Perpustakaan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cetak"
      DragIcon        =   "Petugas.frx":3AB3C4
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Picture         =   "Petugas.frx":3B9130
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      Picture         =   "Petugas.frx":3B9B32
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "Petugas.frx":3B9F74
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "Petugas.frx":3BA2FE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      Picture         =   "Petugas.frx":3BA888
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Picture         =   "Petugas.frx":3BAC12
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   8400
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   8400
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   5520
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "NIP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rspetugas As Recordset
Public x
Dim Pilih

Private Sub Command1_Click()
x = 1
ardin
End Sub

Private Sub Command2_Click()
x = 2
Text1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Command3_Click()
Bersih
End Sub

Private Sub Command4_Click()
x = 3
pesan = MsgBox("Yakin data akan dihapus", vbYesNo + vbQuestion, "information")
If pesan = vbYes Then
    rspetugas.Delete
    Bersih
End If
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command7_Click()
Simpan
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Pilih = File1.Path & "\" & File1.FileName
On Error GoTo kosong
Image1.Picture = LoadPicture(Pilih)

Exit Sub

kosong:
Pilih = Space(100)
pesan = MsgBox("Gambar Kosong", vbOKOnly + vbInformation, "Informasi")



End Sub

Private Sub Form_Activate()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = False


End Sub

Private Sub Form_Load()
Set dbpustaka = OpenDatabase("D:\Perpustakaan\Perpustakaan.mdb")
Set rspetugas = dbpustaka.OpenRecordset("petugas")
End Sub
Function ardin()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True

Text1.SetFocus

End Function
Function Bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = "[Pilih Jenis Kelamin]"
'Text1.SetFocus

End Function
Function Simpan()
If x = 1 Then
rspetugas.AddNew
rspetugas!NIP = Text1
rspetugas!Nama = Text2
rspetugas!Alamat = Text3
rspetugas!jenis_kelamin = Combo1.Text
rspetugas.Update
pesan = MsgBox("Data Telah di Simpan", vbOKOnly + vbInformation, "Information")
Bersih
ElseIf x = 2 Then
rspetugas.Edit
rspetugas!NIP = Text1
rspetugas!Nama = Text2
rspetugas!Alamat = Text3
rspetugas!jenis_kelamin = Combo1.Text
rspetugas.Update
pesan = MsgBox("Data Telah di Ubah", vbOKOnly + vbInformation, "Information")
Bersih
End If
End Function


Private Sub Picture1_Click()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rspetugas.Index = "NIP"
rspetugas.Seek "=", Text1.Text
If rspetugas.NoMatch Then
pesan = MsgBox(" Data Tidak Ada", vbOKOnly + vbInformation, "warning")
Text1.SetFocus
Else
tampil
End If
End If


End Sub

Function tampil()
Text1.Text = rspetugas!NIP
Text2.Text = rspetugas!Nama
Text3.Text = rspetugas!Alamat
Combo1.Text = rspetugas!jenis_kelamin
If rspetugas!lokasi <> Space(100) Then
Image1.Picture = LoadPicture(rssiswa!lokasi)
End If
End Function


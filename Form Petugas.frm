VERSION 5.00
Begin VB.Form Form4 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Petugas"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12750
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form Petugas.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      BackColor       =   &H00000000&
      Connect         =   "Access"
      DatabaseName    =   "G:\FEBRIAN\Perpustakaan\perpustakaan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Width           =   6735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   5880
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   14
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "Form Petugas.frx":50A21
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   7800
      TabIndex        =   10
      Top             =   5640
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   7800
      TabIndex        =   9
      Top             =   5040
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7800
      TabIndex        =   8
      Top             =   4680
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      ItemData        =   "Form Petugas.frx":50E28
      Left            =   4200
      List            =   "Form Petugas.frx":50E32
      TabIndex        =   7
      Text            =   "Pilih Jenis Kelamin"
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   4200
      TabIndex        =   4
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Input Data Petugas"
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
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   8040
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   2400
      TabIndex        =   3
      Top             =   4080
      Width           =   1785
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   2880
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "NIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   2400
      TabIndex        =   0
      Top             =   2400
      Width           =   570
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rspetugas As Recordset
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
resik
End Sub

Private Sub Command4_Click()
X = 3
pesan = MsgBox("Yakin Mau Dihapus?", vbYesNo + vbQuestion, "Information")
If pesan = vbYes Then
    rspetugas.Delete
    Image1.Visible = False
    resik
End If
End Sub

Private Sub Command5_Click()
Form4.Hide
Form9.Show
End Sub

Private Sub Command7_Click()
Save
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
Set rspetugas = dbpustaka.OpenRecordset("petugas")
End Sub
Function papaw()
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
Text1.SetFocus

End Function
Function resik()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = "[PILIH JENIS KELAMIN]"
'Text1.SetFocus
End Function
Function Save()
If X = 1 Then
rspetugas.AddNew
rspetugas!NIP = Text1.Text
rspetugas!Nama = Text2.Text
rspetugas!Alamat = Text3.Text
rspetugas!Jenis_Kelamin = Combo1.Text
rspetugas!lokasi = pilih
rspetugas.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
resik
ElseIf X = 2 Then
rspetugas.Edit
rspetugas!NIP = Text1
rspetugas!Nama = Text2
rspetugas!Alamat = Text3
rspetugas!Jenis_Kelamin = Combo1.Text
rspetugas!lokasi = pilih
rspetugas.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
resik
End If
End Function

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
pilih = File1.Path & "\" & File1.FileName
On eror GoTo Kosong
Image1.Picture = LoadPicture(pilih)
Exit Sub
Kosong:
pilih = Space(100)
pesan = MsgBox("Gambar Kosong", vbOKOnly + vbInformation, "Informasi")

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rspetugas.Index = "NIP"
rspetugas.Seek "=", Text1.Text
If rspetugas.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text2.SetFocus
Else
tampil
End If
End If
End Sub

Function tampil()
Text1.Text = rspetugas!NIP
Text2.Text = rspetugas!Nama
Text3.Text = rspetugas!Alamat
Combo1.Text = rspetugas!Jenis_Kelamin
If rspetugas!lokasi <> Space(100) Then
Image1.Picture = LoadPicture(rspetugas!lokasi)
End If


End Function


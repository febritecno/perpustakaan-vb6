VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FORM SISWA"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "siswa.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2040
      Top             =   7680
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\FEBRIAN\Perpustakaan\perpustakaan.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\FEBRIAN\Perpustakaan\perpustakaan.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "siswa.frx":50A21
      Height          =   1815
      Left            =   7200
      TabIndex        =   19
      Top             =   3840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Command1"
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "nis"
         Caption         =   "nis"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama"
         Caption         =   "nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "kelas"
         Caption         =   "kelas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "jurusan"
         Caption         =   "jurusan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "alamat"
         Caption         =   "alamat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Lokasi"
         Caption         =   "Lokasi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      ItemData        =   "siswa.frx":50A40
      Left            =   3960
      List            =   "siswa.frx":50A74
      TabIndex        =   13
      Text            =   "[PILIH JURUSAN]"
      Top             =   3120
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      ItemData        =   "siswa.frx":50ACE
      Left            =   3960
      List            =   "siswa.frx":50ADB
      TabIndex        =   12
      Text            =   "[PILIH KELAS]"
      Top             =   2640
      Width           =   3735
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1800
      Left            =   5160
      TabIndex        =   11
      Top             =   3840
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1035
      Left            =   1800
      TabIndex        =   10
      Top             =   4320
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   1800
      TabIndex        =   9
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "&Tambah"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Picture         =   "siswa.frx":50AEB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      Picture         =   "siswa.frx":51075
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      Picture         =   "siswa.frx":515FF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Cetak Kartu Anggota"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Picture         =   "siswa.frx":51B89
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "&Ubah"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jurusan"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   18
      Top             =   3240
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kelas"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   17
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   16
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Lengkap"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   15
      Top             =   1560
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIS"
      BeginProperty Font 
         Name            =   "Felix Titling"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbpustaka As Database
Dim rssiswa As Recordset
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
On Error Resume Next
X = 3
pesan = MsgBox("Yakin Mau Dihapus?", vbYesNo + vbQuestion, "Information")
If pesan = vbYes Then
    
    If rssiswa.EOF Then
    MsgBox "Tidak ada data"
    Else
    rssiswa.Delete
    DataGrid1.Refresh
    resik
End If

End If
End Sub

Private Sub Command5_Click()
Form1.Hide
Form9.Show
End Sub

Private Sub Command6_Click()
DataReport1.Show
End Sub

Private Sub Command7_Click()
Save
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
Combo1.Enabled = False
Combo2.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
End Sub

Private Sub Form_Load()
Set dbpustaka = OpenDatabase(App.Path & "\perpustakaan.mdb")
Set rssiswa = dbpustaka.OpenRecordset("siswa")
End Sub
Function papaw()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command6.Enabled = True
Text1.SetFocus

End Function
Function resik()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = "[PILIH KELAS]"
Combo2.Text = "[PILIH JURUSAN]"
Image1.Visible = False
'Text1.SetFocus
End Function
Function Save()
On Error Resume Next
If X = 1 Then
rssiswa.AddNew
rssiswa!nis = Text1.Text
rssiswa!Nama = Text2.Text
rssiswa!Alamat = Text3.Text
rssiswa!kelas = Combo1.Text
rssiswa!jurusan = Combo2.Text
rssiswa!lokasi = pilih
rssiswa.Update
DataGrid1.Refresh
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
resik
ElseIf X = 2 Then
rssiswa.Edit
rssiswa!nis = Text1
rssiswa!Nama = Text2
rssiswa!Alamat = Text3
rssiswa!kelas = Combo1.Text
rssiswa!jurusan = Combo2.Text
rssiswa!lokasi = pilih
rssiswa.Update
pesan = MsgBox("Data Telah Disimpan", vbOKOnly + vbInformation, "Information")
DataGrid1.Refresh
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
On Error Resume Next
Image1.Picture = LoadPicture(pilih)
Exit Sub
Kosong:
pilih = Space(100)
pesan = MsgBox("Gambar Kosong", vbOKOnly + vbInformation, "Informasi")

End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rssiswa.Index = "nis"
rssiswa.Seek "=", Text1.Text
If rssiswa.NoMatch Then
pesan = MsgBox("Data tidak ada", vbOKOnly + vbInformation, "Warning")
Text2.SetFocus
Else
tampil
End If
End If
End Sub

Function tampil()
Text1.Text = rssiswa!nis
Text2.Text = rssiswa!Nama
Text3.Text = rssiswa!Alamat
Combo1.Text = rssiswa!kelas
Combo2.Text = rssiswa!jurusan
If rssiswa!lokasi <> Space(100) Then
Image1.Picture = LoadPicture(rssiswa!lokasi)
End If
End Function




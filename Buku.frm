VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Buku"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5940
   Icon            =   "Buku.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Buku.frx":038A
   ScaleHeight     =   6150
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Batal"
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   1560
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Pengarang"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Penerbit"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Judul Buku"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nomer Buku"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nama Kategori"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nomer Kategori"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

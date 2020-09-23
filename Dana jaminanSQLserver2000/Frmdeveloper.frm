VERSION 5.00
Begin VB.Form Frmdeveloper 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Tabel Developer"
   ClientHeight    =   5010
   ClientLeft      =   4320
   ClientTop       =   3195
   ClientWidth     =   7560
   Icon            =   "Frmdeveloper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin DanaJaminan.vbButton cmdbatal 
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "BATAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":0582
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdkeluar 
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "KELUAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":059E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdhapus 
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "HAPUS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":05BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdubah 
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UBAH"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":05D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdcari 
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "CARI"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":05F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdsimpan 
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "SIMPAN"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":060E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdtambah 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "TAMBAH"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Frmdeveloper.frx":062A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtlokasiproyek 
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox txtpemilik 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtrek 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox txtnpwp 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox txttelp 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtalamat 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox txtnama 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtkd 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      Caption         =   "Lokasi Proyek"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
      Caption         =   "Pemilik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Nomer Rekening"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "NPWP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Telepon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Alamat Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Nama Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Kode Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Frmdeveloper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSdeveloper As New ADODB.Recordset

Private Sub Command5_Click()
Unload Me
MDIDJ.Enabled = True
End Sub

Private Sub CmdBatal_Click()
Call posisiawal
End Sub

Private Sub CmdCari_Click()
Call CariDatabase
Call TxtHidup
cmdubah.Enabled = True
cmdhapus.Enabled = True
End Sub
Public Sub CariDatabase()
    MDIForm1.Enabled = False
    FrmListDeveloper.Show
End Sub

Private Sub cmdhapus_Click()
SQL = "select * from tabeldeveloper where kode_developer='" & Trim(txtkd.Text) & "'"
If RSdeveloper.State > 0 Then RSdeveloper.Close
RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Apa Anda Yakin Mau Hapus Developer Ini?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RSdeveloper.Delete
    MsgBox "data Developer Dihapus.", vbInformation, "Informasi"
    RSdeveloper.Close
End If
TxtKosong
cmdhapus.Enabled = False
End Sub

Private Sub posisiawal()
  Call TxtKosong
  Call TxtMati
  cmdsimpan.Enabled = False
  cmdhapus.Enabled = False
  CmdCari.Enabled = True
  cmdtambah.Enabled = True
  cmdubah.Enabled = False
End Sub

Private Sub CmdKeluar_Click()
Unload Me
MDIForm1.Enabled = True
End Sub

Private Sub cmdsimpan_Click()
If RSdeveloper.State > 0 Then RSdeveloper.Close
    'Save data
    RSdeveloper.Open "tabeldeveloper", cn, adOpenDynamic, adLockOptimistic
    With RSdeveloper
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "data developer disimpan.", vbInformation, "Informasi"
    TxtKosong
End Sub
Private Sub Simpan()
'For saving data
With RSdeveloper
    .Fields("kode_developer") = Trim(txtkd.Text)
    .Fields("nama_developer") = Trim(txtnama.Text)
    .Fields("alamat") = Trim(txtalamat.Text)
    .Fields("telepon") = Trim(txttelp.Text)
    .Fields("NPWP") = Trim(txtnpwp.Text)
    .Fields("pemilik") = Trim(txtpemilik.Text)
    .Fields("lokasi") = Trim(txtlokasiproyek.Text)
    .Fields("no_rekening") = Trim(txtrek.Text)
End With
End Sub

Private Sub cmdtambah_Click()
Call TxtHidup
cmdsimpan.Enabled = True
txtkd.SetFocus
End Sub

Private Sub cmdubah_Click()
'Edit data
    SQL = "select * from tabeldeveloper where kode_developer='" & Trim(txtkd.Text) & "'"
    If RSdeveloper.State > 0 Then RSdeveloper.Close
    RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
    With RSdeveloper
        Simpan
        .Update
        .Close
    End With
    MsgBox "Developer Sudah Diubah.", vbInformation, "Information"
    TxtKosong
    cmdubah.Enabled = False
End Sub

Private Sub TxtKosong()
  txtkd.Text = ""
  txtnama.Text = ""
  txtalamat.Text = ""
  txttelp.Text = ""
  txtnpwp.Text = ""
  txtpemilik.Text = ""
  txtlokasiproyek.Text = ""
  txtrek.Text = ""
End Sub
Private Sub TxtMati()
  txtkd.Enabled = False
  txtnama.Enabled = False
  txtalamat.Enabled = False
  txttelp.Enabled = False
  txtnpwp.Enabled = False
  txtpemilik.Enabled = False
  txtlokasiproyek.Enabled = False
  txtrek.Enabled = False
End Sub

Private Sub TxtHidup()
  txtkd.Enabled = True
  txtnama.Enabled = True
  txtalamat.Enabled = True
  txttelp.Enabled = True
  txtnpwp.Enabled = True
  txtpemilik.Enabled = True
  txtlokasiproyek.Enabled = True
  txtrek.Enabled = True
End Sub

Private Sub Form_Load()
  Me.Width = 7650
  Me.Height = 5415
  Me.Top = 2835
  Me.Left = 4275
  Bukakoneksi
  RSdeveloper.CursorLocation = adUseClient
  Call posisiawal
End Sub

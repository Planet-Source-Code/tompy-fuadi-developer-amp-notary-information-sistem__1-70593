VERSION 5.00
Begin VB.Form Frmnotaris 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Tabel Notaris"
   ClientHeight    =   4425
   ClientLeft      =   4515
   ClientTop       =   3195
   ClientWidth     =   7845
   Icon            =   "Frmnotaris.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin DanaJaminan.vbButton cmdkeluar 
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":0582
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
      Left            =   5520
      TabIndex        =   16
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":059E
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
      Left            =   4440
      TabIndex        =   15
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":05BA
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
      Left            =   3360
      TabIndex        =   14
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":05D6
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
      Left            =   1200
      TabIndex        =   13
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":05F2
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
      Left            =   240
      TabIndex        =   12
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":060E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtrek 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
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
   Begin VB.TextBox txtkdn 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin DanaJaminan.vbButton cmdbatal 
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   3600
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
      MICON           =   "Frmnotaris.frx":062A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   11
      Top             =   2760
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Alamat Notaris"
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
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Nama Notaris"
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
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Kode Notaris"
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
      Width           =   1335
   End
End
Attribute VB_Name = "Frmnotaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSnotaris As New ADODB.Recordset

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
    FrmListNotaris.Show
End Sub

Private Sub cmdhapus_Click()
SQL = "select * from tabelnotaris where kode_notaris='" & Trim(txtkdn.Text) & "'"
If RSnotaris.State > 0 Then RSnotaris.Close
RSnotaris.Open SQL, cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Apa Anda Yakin Mau Hapus notaris Ini?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RSnotaris.Delete
    MsgBox "data notaris Dihapus.", vbInformation, "Informasi"
    RSnotaris.Close
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
If RSnotaris.State > 0 Then RSnotaris.Close
    'Save data
    RSnotaris.Open "tabelnotaris", cn, adOpenDynamic, adLockOptimistic
    With RSnotaris
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "data notaris disimpan.", vbInformation, "Informasi"
    TxtKosong
End Sub
Private Sub Simpan()
'For saving data
With RSnotaris
    .Fields("kode_notaris") = Trim(txtkdn.Text)
    .Fields("nama_notaris") = Trim(txtnama.Text)
    .Fields("alamat") = Trim(txtalamat.Text)
    .Fields("telepon") = Trim(txttelp.Text)
    .Fields("NPWP") = Trim(txtnpwp.Text)
    .Fields("no_rekening") = Trim(txtrek.Text)
End With
End Sub

Private Sub cmdtambah_Click()
Call TxtHidup
cmdsimpan.Enabled = True
txtkdn.SetFocus
End Sub

Private Sub cmdubah_Click()
'Edit data
    SQL = "select * from tabelnotaris where kode_notaris='" & Trim(txtkdn.Text) & "'"
    If RSnotaris.State > 0 Then RSnotaris.Close
    RSnotaris.Open SQL, cn, adOpenDynamic, adLockOptimistic
    With RSnotaris
        Simpan
        .Update
        .Close
    End With
    MsgBox "notaris Sudah Diubah.", vbInformation, "Information"
    TxtKosong
    cmdubah.Enabled = False
End Sub

Private Sub TxtKosong()
  txtkdn.Text = ""
  txtnama.Text = ""
  txtalamat.Text = ""
  txttelp.Text = ""
  txtnpwp.Text = ""
  txtrek.Text = ""
End Sub
Private Sub TxtMati()
  txtkdn.Enabled = False
  txtnama.Enabled = False
  txtalamat.Enabled = False
  txttelp.Enabled = False
  txtnpwp.Enabled = False
  txtrek.Enabled = False
End Sub

Private Sub TxtHidup()
  txtkdn.Enabled = True
  txtnama.Enabled = True
  txtalamat.Enabled = True
  txttelp.Enabled = True
  txtnpwp.Enabled = True
  txtrek.Enabled = True
End Sub

Private Sub Form_Load()
  Me.Width = 7935
  Me.Height = 4830
  Me.Top = 2835
  Me.Left = 4470
  Bukakoneksi
  RSnotaris.CursorLocation = adUseClient
  Call posisiawal
End Sub

VERSION 5.00
Begin VB.Form FrmKredit 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Input Tabel Kredit"
   ClientHeight    =   1530
   ClientLeft      =   5295
   ClientTop       =   3960
   ClientWidth     =   4845
   Icon            =   "FrmKredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin DanaJaminan.vbButton cmdbaru 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "BARU"
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
      MICON           =   "FrmKredit.frx":0582
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
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "FrmKredit.frx":059E
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
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "FrmKredit.frx":05BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtkredit 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "JENIS KREDIT"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmKredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSkredit As New ADODB.Recordset

Private Sub cmdbaru_Click()
Call TxtHidup
cmdtambah.Enabled = True
End Sub

Private Sub CmdKeluar_Click()
Unload Me
MDIForm1.Enabled = True
End Sub

Private Sub cmdtambah_Click()
If RSkredit.State > 0 Then RSkredit.Close
    'Save data
    RSkredit.Open "tabelkredit", cn, adOpenDynamic, adLockOptimistic
    With RSkredit
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "kredit data disimpan.", vbInformation, "Informasi"
    TxtKosong
End Sub
Private Sub Simpan()
'For saving data
With RSkredit
    .Fields("jeniskredit") = Trim(txtkredit.Text)
End With
End Sub

Private Sub TxtKosong()
  txtkredit.Text = ""
  cmdtambah.Enabled = False
  cmdbaru.Enabled = True
  Call TxtMati
End Sub
Private Sub TxtMati()
  txtkredit.Enabled = False
  cmdtambah.Enabled = False
End Sub

Private Sub TxtHidup()
  txtkredit.Enabled = True
  cmdbaru.Enabled = False
  txtkredit.SetFocus
End Sub

Private Sub Form_Load()
  Me.Width = 4935
  Me.Height = 1860
  Me.Top = 4515
  Me.Left = 5790
Bukakoneksi
RSkredit.CursorLocation = adUseClient
Call TxtMati
End Sub

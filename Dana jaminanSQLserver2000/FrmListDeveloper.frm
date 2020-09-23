VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListDeveloper 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Notaris"
   ClientHeight    =   6735
   ClientLeft      =   2760
   ClientTop       =   2235
   ClientWidth     =   10605
   ControlBox      =   0   'False
   Icon            =   "FrmListDeveloper.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGridbase 
      Bindings        =   "FrmListDeveloper.frx":0582
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Data Developer"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin DanaJaminan.vbButton CmdCari 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
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
      MICON           =   "FrmListDeveloper.frx":059C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton CmdBatal 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
      MICON           =   "FrmListDeveloper.frx":05B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton CmdOk 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "OK"
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
      MICON           =   "FrmListDeveloper.frx":05D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtkode 
      Height          =   315
      Left            =   2625
      MaxLength       =   25
      TabIndex        =   0
      Top             =   6240
      Width           =   2355
   End
   Begin VB.Label CARI 
      BackStyle       =   0  'Transparent
      Caption         =   "&CARI KODE DEVELOPER  :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmListDeveloper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSdeveloper As New ADODB.Recordset

Private Sub CmdBatal_Click()
  MDIForm1.Enabled = True
  Unload Me
End Sub

Private Sub CmdCari_Click()
If Trim(txtkode) <> "" Then
    SQL = "select * from tabeldeveloper where kode_developer='" & txtkode.Text & "'"
    Set RSdeveloper = cn.Execute(SQL)
    Set DataGridbase.DataSource = RSdeveloper
End If
End Sub

Private Sub CmdOk_Click()
  Call DataOk
End Sub

Private Sub DataOk()
      Frmdeveloper.txtkd = RSdeveloper!kode_developer
      Frmdeveloper.txtnama = RSdeveloper!nama_developer
      Frmdeveloper.txtalamat = RSdeveloper!alamat
      Frmdeveloper.txttelp = RSdeveloper!telepon
      Frmdeveloper.txtnpwp = RSdeveloper!NPWP
      Frmdeveloper.txtpemilik = RSdeveloper!pemilik
      Frmdeveloper.txtlokasiproyek = RSdeveloper!lokasi
      Frmdeveloper.txtrek = RSdeveloper!no_rekening
      Call Frmdeveloper.CariDatabase
    MDIForm1.Enabled = True
    Unload Me
End Sub

Private Sub DataGridbase_Click()
Call DataOk
End Sub
Private Sub Form_Activate()
txtkode.SetFocus
End Sub

Private Sub Form_Load()
Bukakoneksi
cn.CursorLocation = adUseClient
SQL = "select * from tabeldeveloper order by kode_developer"
Set RSdeveloper = cn.Execute(SQL)
Set DataGridbase.DataSource = RSdeveloper
End Sub

Private Sub Txtkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdCari_Click
End If
End Sub




VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmListBase1 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Information"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   915
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "FrmListBase2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGridbase 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
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
      Left            =   5400
      TabIndex        =   3
      Top             =   6000
      Width           =   990
      _ExtentX        =   1746
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
      BCOLO           =   65535
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmListBase2.frx":0582
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
      Width           =   2055
      _ExtentX        =   3625
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
      BCOLO           =   65535
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmListBase2.frx":059E
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
      Left            =   720
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
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
      BCOLO           =   65535
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmListBase2.frx":05BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Txtkode 
      Height          =   315
      Left            =   2865
      MaxLength       =   25
      TabIndex        =   0
      Top             =   6120
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI NOMER DEBITUR   :"
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
      TabIndex        =   5
      Top             =   6240
      Width           =   2535
   End
End
Attribute VB_Name = "FrmListBase1"
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
If Trim(Txtkode) <> "" Then
    SQL = "select * from datadeveloper where nomerdebitur='" & Txtkode.Text & "'"
    Set RSdeveloper = cn.Execute(SQL)
    Set DataGridbase.DataSource = RSdeveloper
End If
End Sub

Private Sub CmdOk_Click()
  Call DataOk
End Sub

Private Sub DataOk()
    Formcair.txtnomer = RSdeveloper!Nomer
    Formcair.txtno = RSdeveloper!nomerdebitur
    Formcair.txtnama = RSdeveloper!namadebitur
    Formcair.cbjenis = RSdeveloper!jeniskredit
    Formcair.txtdeveloper = RSdeveloper!developer
    Formcair.txtmak = RSdeveloper!maksimalkredit
    Formcair.txtimb = RSdeveloper!jaminanimb
    Formcair.txtser = RSdeveloper!jaminansertifikat
    Formcair.txtlistrik = RSdeveloper!jaminanlistrik
    Formcair.txtjalan = RSdeveloper!jaminanbestekjalan
    Formcair.txtair = RSdeveloper!jaminanbestekair
      Call Formcair.CariDatabase
    MDIForm1.Enabled = True
    Unload Me
End Sub

Private Sub DataGridbase_Click()
Call DataOk
End Sub
Private Sub Form_Activate()
Txtkode.SetFocus
End Sub

Private Sub Form_Load()
Bukakoneksi
cn.CursorLocation = adUseClient
SQL = "select * from datadeveloper order by nomer"
Set RSdeveloper = cn.Execute(SQL)
Set DataGridbase.DataSource = RSdeveloper
End Sub

Private Sub Txtkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdCari_Click
End If
End Sub






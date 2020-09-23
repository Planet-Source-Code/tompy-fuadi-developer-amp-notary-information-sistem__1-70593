VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmdev 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jaminan Developer"
   ClientHeight    =   1365
   ClientLeft      =   5835
   ClientTop       =   4365
   ClientWidth     =   3960
   Icon            =   "frmdev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtnamadev 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtdev 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin DanaJaminan.vbButton cmdcetak 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "CETAK"
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
      MICON           =   "frmdev.frx":0582
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Developer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmdev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSdeveloper As New ADODB.Recordset

Private Sub cmdcetak_Click()
Cr.ReportFileName = App.Path & "\Danajaminandeveloper.rpt"
Cr.WindowTitle = "Dana Jaminan Report"
Cr.SelectionFormula = "({datadeveloper.developer} = '" & Me.txtdev.Text & "')"
Cr.WindowShowRefreshBtn = True
Cr.Destination = crptToWindow
Cr.Action = 1
End Sub

Private Sub Form_Load()
  Me.Width = 3930
  Me.Height = 1860
  Me.Top = 4515
  Me.Left = 5790
  Bukakoneksi
RSdeveloper.CursorLocation = adUseClient
End Sub

Private Sub txtdev_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If txtdev <> "" Then
      Call CariDatadev
    Else
      FrmListdev1.Show
    End If
  End If
End Sub
Public Sub CariDatadev()
  MDIForm1.Enabled = False
  FrmListdev1.Show
End Sub

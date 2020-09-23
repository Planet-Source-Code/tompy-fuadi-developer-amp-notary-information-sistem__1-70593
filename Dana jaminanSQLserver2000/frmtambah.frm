VERSION 5.00
Begin VB.Form frmtambah 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tambah User"
   ClientHeight    =   1455
   ClientLeft      =   5835
   ClientTop       =   4875
   ClientWidth     =   3840
   Icon            =   "frmtambah.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin DanaJaminan.vbButton cmdkeluar 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmtambah.frx":0442
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
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmtambah.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdbaru 
      Height          =   375
      Left            =   120
      TabIndex        =   2
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
      MICON           =   "frmtambah.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name    :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmtambah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RStambah As New ADODB.Recordset

Private Sub cmdbaru_Click()
MsgBox "Insert your username and password into the textboxes and click on add."
cmdbaru.Enabled = False
cmdtambah.Enabled = True
Call isi
txtUsername.SetFocus 'sets the cursor to txtusername.text
End Sub

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdtambah_Click()
If RStambah.State > 0 Then RStambah.Close
    'Save data
    RStambah.Open "users", cn, adOpenDynamic, adLockOptimistic
    With RStambah
        .AddNew
        Simpan
        .Update
        .Close
    End With
    MsgBox "user data disimpan.", vbInformation, "Informasi"
    txtPass.Text = ""
    txtUsername.Text = ""
    MsgBox "Complete!"
    cmdbaru.Enabled = True
    cmdtambah.Enabled = False
    txtUsername.SetFocus 'sets the cursor to txtusername.text
End Sub

Private Sub Simpan()
'For saving data
With RStambah
    .Fields("Username") = Trim(txtUsername.Text)
    .Fields("Password") = Trim(txtPass.Text)
End With
End Sub

Private Sub Form_Load()
  Me.Width = 3930
  Me.Height = 1860
  Me.Top = 4515
  Me.Left = 5790
Bukakoneksi
RStambah.CursorLocation = adUseClient
Call kosong
End Sub

Private Sub kosong()
txtUsername.Enabled = False
txtPass.Enabled = False
End Sub

Private Sub isi()
txtUsername.Enabled = True
txtPass.Enabled = True
End Sub

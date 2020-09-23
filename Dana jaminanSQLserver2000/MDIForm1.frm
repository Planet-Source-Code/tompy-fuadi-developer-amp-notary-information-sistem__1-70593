VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MENU DANA JAMINAN"
   ClientHeight    =   4485
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   6555
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0582
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4110
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "28/05/2008"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "1:33 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "PROGRAM DANA JAMINAN DEVELOPER & NOTARIS"
            TextSave        =   "PROGRAM DANA JAMINAN DEVELOPER & NOTARIS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Versi 1.0"
            TextSave        =   "Versi 1.0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport Cr2 
      Left            =   6840
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   1320
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":695A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6D6F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":95396
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BD038
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C077A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":E841C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1100BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":137D60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":15FA02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1876A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1AF346
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1482
      ButtonWidth     =   2805
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INPUT DATA KREDIT"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PENCAIRAN"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INPUT KREDIT"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INPUT DEVELOPER"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INPUT NOTARIS"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MONITORING"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TAMBAH USER"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXIT"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu util 
      Caption         =   "&Utility"
      Begin VB.Menu inputdev 
         Caption         =   "1. Input Tabel Developer"
      End
      Begin VB.Menu inputnot 
         Caption         =   "2. Input Tabel Notaris"
      End
      Begin VB.Menu inputkredit 
         Caption         =   "3. Input Kredit"
      End
      Begin VB.Menu user 
         Caption         =   "4. Tambah User"
      End
   End
   Begin VB.Menu input 
      Caption         =   "&Input Data"
      Begin VB.Menu Inputmaster 
         Caption         =   "&INPUT MASTER DANA JAMINAN"
      End
   End
   Begin VB.Menu cair 
      Caption         =   "&Pencairan"
      Begin VB.Menu cairmaster 
         Caption         =   "&PENCAIRAN MASTER DANA JAMINAN"
      End
   End
   Begin VB.Menu report 
      Caption         =   "&Report"
      Begin VB.Menu kdjd 
         Caption         =   "&JAMINAN DEVELOPER"
         Begin VB.Menu rinciandev 
            Caption         =   "&KARTU DANA JAMINAN DEVELOPER (Rincian)"
         End
         Begin VB.Menu rekapdev 
            Caption         =   "REKAP DANA JAMINAN DEVELOPER"
         End
      End
      Begin VB.Menu kdjn 
         Caption         =   "&JAMINAN NOTARIS"
         Begin VB.Menu rinciannot 
            Caption         =   "&KARTU DANA JAMINAN NOTARIS ( Rincian )"
         End
         Begin VB.Menu rekapnot 
            Caption         =   "REKAP DANA JAMINAN NOTARIS"
         End
      End
      Begin VB.Menu RBND 
         Caption         =   "&RINCIAN DANA NOTARIS YANG BELUM DIBAYAR"
         Begin VB.Menu RBN 
            Caption         =   "RINCIAN BIAYA NOTARIS"
         End
         Begin VB.Menu BN 
            Caption         =   "&ALL BIAYA NOTARIS"
         End
         Begin VB.Menu RBA 
            Caption         =   "RINCIAN BIAYA APHT"
         End
         Begin VB.Menu BAPHT 
            Caption         =   "&ALL BIAYA APHT"
         End
      End
      Begin VB.Menu Register 
         Caption         =   "&REGISTER REALISASI DAN MONITORING PER-TANGGAL"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BAPHT_Click()
cr2.ReportFileName = App.Path & "\REPORT ALL RINCIAN BIAYA APHT YANG BELUM DIBAYAR.rpt"
cr2.WindowTitle = "REPORT RINCIAN BIAYA APHT YANG BELUM DIBAYAR"
cr2.WindowShowRefreshBtn = True
cr2.Destination = crptToWindow
cr2.Action = 1
End Sub

Private Sub BN_Click()
Cr1.ReportFileName = App.Path & "\REPORT ALL RINCIAN BIAYA NOTARIS YANG BELUM DIBAYAR.rpt"
Cr1.WindowTitle = "REPORT RINCIAN BIAYA NOTARIS YANG BELUM DIBAYAR"
Cr1.WindowShowRefreshBtn = True
Cr1.Destination = crptToWindow
Cr1.Action = 1
End Sub

Private Sub cairmaster_Click()
Formcair.Show
End Sub

Private Sub inputdev_Click()
Frmdeveloper.Show
End Sub

Private Sub inputkredit_Click()
FrmKredit.Show
End Sub

Private Sub Inputmaster_Click()
Formlocal.Show
End Sub

Private Sub inputnot_Click()
Frmnotaris.Show
End Sub

Private Sub keluar_Click()
End
End Sub

Private Sub reportall_Click()
FrmListBase.Show
End Sub

Private Sub RBA_Click()
frmrinciannot2.Show
End Sub

Private Sub RBN_Click()
frmrinciannot1.Show
End Sub

Private Sub Register_Click()
Cr.ReportFileName = App.Path & "\REGISTER REALISASI DAN MONITORING.rpt"
Cr.WindowTitle = "Dana Jaminan Register Report"
Cr.WindowShowRefreshBtn = True
Cr.Destination = crptToWindow
Cr.Action = 1
End Sub

Private Sub rekapdev_Click()
Formrekapdeveloper.Show
End Sub

Private Sub rekapnot_Click()
Formrekapnotaris.Show
End Sub

Private Sub rinciandev_Click()
frmdev.Show
End Sub

Private Sub rinciannot_Click()
frmnot.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1:
Formlocal.Show
Case 2:
Formcair.Show
Case 3:
FrmKredit.Show
Case 4:
Frmdeveloper.Show
Case 5:
Frmnotaris.Show
Case 6:
Cr.ReportFileName = App.Path & "\REGISTER REALISASI DAN MONITORING.rpt"
Cr.WindowTitle = "Dana Jaminan Register Report"
Cr.WindowShowRefreshBtn = True
Cr.Destination = crptToWindow
Cr.Action = 1
Case 7:
frmtambah.Show
Case 8:
Unload Me
End Select
End Sub

Private Sub user_Click()
frmtambah.Show
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formcair 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Pencairan"
   ClientHeight    =   6975
   ClientLeft      =   3570
   ClientTop       =   2115
   ClientWidth     =   8310
   Icon            =   "Formoverseas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtno 
      Height          =   285
      Left            =   2520
      TabIndex        =   25
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtimb 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtser 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   19
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtlistrik 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   20
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtjalan 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox txtair 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   22
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtmak 
      Height          =   285
      Left            =   3120
      TabIndex        =   39
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtnama 
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox cbjenis 
      Height          =   315
      Left            =   3120
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtstatus 
      Height          =   285
      Left            =   2040
      TabIndex        =   26
      Text            =   "D"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtstatus2 
      Height          =   285
      Left            =   2040
      TabIndex        =   31
      Text            =   "D"
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtbiayaapht 
      Height          =   285
      Left            =   2520
      TabIndex        =   24
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtbiayan 
      Height          =   285
      Left            =   2520
      TabIndex        =   23
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtnama2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtno2 
      Height          =   285
      Left            =   2520
      TabIndex        =   35
      Top             =   4560
      Width           =   1455
   End
   Begin DanaJaminan.vbButton cmdkeluar 
      Height          =   375
      Left            =   3600
      TabIndex        =   33
      Top             =   6120
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "KELUAR"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":0582
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":05AE
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdbatal2 
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Top             =   5640
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "BATAL"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":05CC
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":05F8
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdubah2 
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "DEBET"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":0616
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":0642
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdcari2 
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "CARI"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":0660
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":068C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdsimpan2 
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "SIMPAN"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":06AA
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":06D6
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdbatal 
      Height          =   375
      Left            =   6840
      TabIndex        =   34
      Top             =   2520
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "BATAL"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":06F4
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":0720
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdubah 
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "DEBET"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":073E
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":076A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdcari 
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "CARI"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":0788
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":07B4
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdsimpan 
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      btype           =   3
      tx              =   "SIMPAN"
      enab            =   -1  'True
      font            =   "Formoverseas.frx":07D2
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      bcolo           =   255
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Formoverseas.frx":07FE
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "PENCAIRAN DANA JAMINAN NOTARIS :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   40
      Top             =   3840
      Width           =   8055
      Begin Crystal.CrystalReport cr 
         Left            =   5400
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin DanaJaminan.vbButton vbButton1 
         Height          =   375
         Left            =   6720
         TabIndex        =   62
         Top             =   2280
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         btype           =   3
         tx              =   "PRINT"
         enab            =   -1  'True
         font            =   "Formoverseas.frx":081C
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   65535
         bcolo           =   255
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Formoverseas.frx":0848
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.TextBox txtnotaris 
         Height          =   285
         Left            =   1920
         TabIndex        =   61
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtuapht 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   59
         Text            =   "0"
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtunotaris 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   58
         Text            =   "0"
         Top             =   1440
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20578307
         CurrentDate     =   30196
      End
      Begin VB.TextBox txtdbiayaapht 
         Height          =   285
         Left            =   4320
         TabIndex        =   15
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtdbiayan 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtnomer2 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtstatusapht 
         Height          =   285
         Left            =   1920
         TabIndex        =   32
         Text            =   "D"
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Transaksi  "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Debitur"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   52
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C00000&
         Caption         =   "&Tanggal"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   43
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C00000&
         Caption         =   "&Biaya Notaris"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   42
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C00000&
         Caption         =   "Biaya APHT"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   41
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "PENCAIRAN DANA JAMINAN DEVELOPER :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtnomer 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   30196
      End
      Begin VB.TextBox txtdser 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtdair 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtdjalan 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtdlistrik 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtdimb 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin Crystal.CrystalReport cr2 
         Left            =   960
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin DanaJaminan.vbButton vbButton2 
         Height          =   375
         Left            =   6720
         TabIndex        =   63
         Top             =   2880
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         btype           =   3
         tx              =   "PRINT"
         enab            =   -1  'True
         font            =   "Formoverseas.frx":0866
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   65535
         bcolo           =   255
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "Formoverseas.frx":0892
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.TextBox txtdeveloper 
         Height          =   285
         Left            =   2400
         TabIndex        =   60
         Text            =   "developer"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtuair 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   57
         Text            =   "0"
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox txtujalan 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   56
         Text            =   "0"
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox txtulistrik 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   55
         Text            =   "0"
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   54
         Text            =   "0"
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtuimb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6240
         TabIndex        =   53
         Text            =   "0"
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtstatusser 
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Text            =   "D"
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatuslistrik 
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Text            =   "D"
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatusjalan 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Text            =   "D"
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatusair 
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Text            =   "D"
         Top             =   2880
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbtgl 
         BackStyle       =   0  'Transparent
         Caption         =   "##/##/####"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6720
         TabIndex        =   65
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Debitur"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   64
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Transaksi  "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   51
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C00000&
         Caption         =   "&Tanggal    "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   50
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C00000&
         Caption         =   "&Jaminan IMB   "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   49
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C00000&
         Caption         =   "&Jaminan Sertifikat            "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   48
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C00000&
         Caption         =   "&Jaminan Listrik"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   47
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C00000&
         Caption         =   "&Jaminan Bestek Jalan     "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   46
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C00000&
         Caption         =   "&Jaminan Bestek Air "
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   45
         Top             =   2880
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Formcair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSdeveloper As New ADODB.Recordset
Dim zTgl As Date

Private Sub CmdBatal_Click()
Call PosisiAwal
End Sub

Private Sub cmdbatal2_Click()
Call PosisiAwal2
End Sub

Private Sub CmdCari_Click()
Call CariDatabase
Call TxtMati
cmdubah.Enabled = True
End Sub

Public Sub CariDatabase()
  MDIForm1.Enabled = False
  FrmListBase1.Show
End Sub

Private Sub cmdcari2_Click()
Call CariDatabase2
Call TxtMati2
cmdubah.Enabled = True
End Sub

Public Sub CariDatabase2()
 MDIForm1.Enabled = False
 FrmListBaseOverseas1.Show
End Sub

Private Sub PosisiAwal()
  Call TxtKosong
  Call TxtMati
  cmdsimpan.Enabled = False
  CmdCari.Enabled = True
  cmdubah.Enabled = True
End Sub

Private Sub PosisiAwal2()
  Call TxtKosong2
  Call TxtMati2
  cmdsimpan2.Enabled = False
  cmdcari2.Enabled = True
  cmdubah2.Enabled = True
End Sub

Private Sub CmdKeluar_Click()
Unload Me
MDIForm1.Enabled = True
End Sub

Private Sub cmdsimpan_Click()
If RSdeveloper.State > 0 Then RSdeveloper.Close
    'Save data
    RSdeveloper.Open "datadeveloper", cn, adOpenDynamic, adLockOptimistic
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
    .Fields("Nomer") = Trim(txtnomer.Text)
    .Fields("nomerdebitur") = Trim(txtno.Text)
    .Fields("namadebitur") = Trim(txtnama.Text)
    .Fields("jeniskredit") = Trim(cbjenis.Text)
    .Fields("developer") = Trim(txtdeveloper.Text)
    .Fields("tanggalposting") = Trim(DTPicker1.Value)
    .Fields("maksimalkredit") = Trim(txtmak.Text)
    .Fields("jaminanimb") = Trim(txtimb.Text)
    .Fields("debetjaminanimb") = Trim(-txtdimb.Text)
    .Fields("statusimb") = Trim(txtstatus.Text)
    .Fields("unitimb") = Trim(txtuimb.Text)
    .Fields("jaminansertifikat") = Trim(txtser.Text)
    .Fields("debetjaminansertifikat") = Trim(-txtdser.Text)
    .Fields("statussertifikat") = Trim(txtstatusser.Text)
    .Fields("unitsertifikat") = Trim(TxtUser.Text)
    .Fields("jaminanlistrik") = Trim(txtlistrik.Text)
    .Fields("debetjaminanlistrik") = Trim(-txtdlistrik.Text)
    .Fields("statuslistrik") = Trim(txtstatuslistrik.Text)
    .Fields("unitlistrik") = Trim(txtulistrik.Text)
    .Fields("jaminanbestekjalan") = Trim(txtjalan.Text)
    .Fields("debetjaminanbestekjalan") = Trim(-txtdjalan.Text)
    .Fields("statusjalan") = Trim(txtstatusjalan.Text)
    .Fields("unitjalan") = Trim(txtujalan.Text)
    .Fields("jaminanbestekair") = Trim(txtair.Text)
    .Fields("debetjaminanbestekair") = Trim(-txtdair.Text)
    .Fields("statusair") = Trim(txtstatusair.Text)
    .Fields("unitair") = Trim(txtuair.Text)
End With
End Sub

Private Sub cmdsimpan2_Click()
If RSdeveloper.State > 0 Then RSdeveloper.Close
    'Save data
    RSdeveloper.Open "datanotaris", cn, adOpenDynamic, adLockOptimistic
    With RSdeveloper
        .AddNew
        Simpan2
        .Update
        .Close
    End With
    MsgBox "data notaris disimpan.", vbInformation, "Informasi"
    TxtKosong2
End Sub
Private Sub Simpan2()
'For saving data
With RSdeveloper
    .Fields("Nomer") = Trim(txtnomer2.Text)
    .Fields("nomerdebitur") = Trim(txtno2.Text)
    .Fields("namadebitur") = Trim(txtnama2.Text)
    .Fields("tanggalposting") = Trim(DTPicker2.Value)
    .Fields("notaris") = Trim(txtnotaris.Text)
    .Fields("biayanotaris") = Trim(txtbiayan.Text)
    .Fields("debetbiayanotaris") = Trim(-txtdbiayan.Text)
    .Fields("statusnotaris") = Trim(txtstatus2.Text)
    .Fields("unitnotaris") = Trim(txtunotaris.Text)
    .Fields("biayaapht") = Trim(txtbiayaapht.Text)
    .Fields("debetbiayaapht") = Trim(-txtdbiayaapht.Text)
    .Fields("statusapht") = Trim(txtstatusapht.Text)
    .Fields("unitapht") = Trim(txtuapht.Text)
End With
End Sub

Private Sub TxtKosong()
     txtnomer.Text = ""
     txtno.Text = ""
     txtnama.Text = ""
     cbjenis.ListIndex = -1
     txtdeveloper.Text = ""
     txtmak.Text = ""
     txtimb.Text = ""
     txtdimb.Text = ""
     txtuimb.Text = "0"
     txtser.Text = ""
     txtdser.Text = ""
     TxtUser.Text = "0"
     txtlistrik.Text = ""
     txtdlistrik.Text = ""
     txtulistrik.Text = "0"
     txtjalan.Text = ""
     txtdjalan.Text = ""
     txtujalan.Text = "0"
     txtair.Text = ""
     txtdair.Text = ""
     txtuair.Text = "0"
End Sub
Private Sub TxtMati()
  txtno.Enabled = False
  txtnama.Enabled = False
  cbjenis.Enabled = False
  txtdeveloper.Enabled = False
  txtmak.Enabled = False
  txtimb.Enabled = False
  txtdimb.Enabled = False
  txtser.Enabled = False
  txtdser.Enabled = False
  txtlistrik.Enabled = False
  txtdlistrik.Enabled = False
  txtjalan.Enabled = False
  txtdjalan.Enabled = False
  txtair.Enabled = False
  txtdair.Enabled = False
End Sub
Private Sub TxtHidup()
  txtno.Enabled = True
  txtnama.Enabled = True
  cbjenis.Enabled = True
  txtdeveloper.Enabled = True
  txtmak.Enabled = True
  txtimb.Enabled = True
  txtdimb.Enabled = True
  txtser.Enabled = True
  txtdser.Enabled = True
  txtlistrik.Enabled = True
  txtdlistrik.Enabled = True
  txtjalan.Enabled = True
  txtdjalan.Enabled = True
  txtair.Enabled = True
  txtdair.Enabled = True
End Sub

Private Sub cmdubah_Click()
  CmdCari.Enabled = True
  cmdubah.Enabled = False
  cmdsimpan.Enabled = True
  Call TxtHidup
  If txtimb.Text = "0" Then txtuimb.Text = "0"
  If txtser.Text = "0" Then TxtUser.Text = "0"
  If txtlistrik.Text = "0" Then txtulistrik.Text = "0"
  If txtjalan.Text = "0" Then txtujalan.Text = "0"
  If txtair.Text = "0" Then txtuair.Text = "0"
End Sub

Private Sub cmdubah2_Click()
  cmdcari2.Enabled = True
  cmdubah2.Enabled = False
  cmdsimpan2.Enabled = True
  Call TxtHidup2
  If txtbiayan.Text = "0" Then txtunotaris.Text = "0"
  If txtbiayaapht.Text = "0" Then txtuapht.Text = "0"
End Sub

Private Sub Form_Load()
  Me.Width = 8400
  Me.Height = 7380
  Me.Top = 1755
  Me.Left = 3525
DTPicker1.Value = Date
DTPicker2.Value = Date
lbtgl.Caption = Format(Date, "dd/mm/yyyy")
  zTgl = Format(Date, "dd/mm/yyyy")
Bukakoneksi
cn.CursorLocation = adUseClient
      CmdBatal_Click
      cmdbatal2_Click
End Sub

Private Sub TxtKosong2()
     txtnomer2.Text = ""
     txtno2.Text = ""
     txtnama2.Text = ""
     txtnotaris.Text = ""
     txtbiayan.Text = ""
     txtdbiayan.Text = ""
     txtunotaris.Text = "0"
     txtbiayaapht.Text = ""
     txtdbiayaapht.Text = ""
     txtuapht.Text = "0"
End Sub
Private Sub TxtMati2()
  txtno2.Enabled = False
  txtnama2.Enabled = False
  txtnotaris.Enabled = False
  txtbiayan.Enabled = False
  txtdbiayan.Enabled = False
  txtbiayaapht.Enabled = False
  txtdbiayaapht.Enabled = False
End Sub

Private Sub TxtHidup2()
  txtno2.Enabled = True
  txtnama2.Enabled = True
  txtnotaris.Enabled = True
  txtbiayan.Enabled = True
  txtdbiayan.Enabled = True
  txtbiayaapht.Enabled = True
  txtdbiayaapht.Enabled = True
  txtnomer2.SetFocus
End Sub

Private Sub txtair_Change()
If txtair.Text = "0" Then txtuair.Text = "-1"
End Sub

Private Sub txtbiayaapht_Change()
If txtbiayaapht.Text = "0" Then txtuapht.Text = "-1"
End Sub

Private Sub txtbiayan_Change()
If txtbiayan.Text = "0" Then txtunotaris.Text = "-1"
End Sub

Private Sub txtdair_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtair.Text = txtair.Text + (-txtdair.Text)
If KeyCode = 13 Then cmdsimpan.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtdbiayaapht_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtbiayaapht.Text = txtbiayaapht.Text + (-txtdbiayaapht.Text)
If KeyCode = 13 Then cmdsimpan2.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtdbiayan_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtbiayan.Text = txtbiayan.Text + (-txtdbiayan.Text)
If KeyCode = 13 Then txtdbiayaapht.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub
Private Sub txtdeveloper_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If txtdeveloper <> "" Then
      Call CariDatadev
    Else
      FrmListdev.Show
    End If
  End If
End Sub
Public Sub CariDatadev()
  Datadeveloper.Recordset.FindFirst "kode_developer = '" & txtdeveloper.Text & "'"
  If Not Datadeveloper.Recordset.NoMatch Then
    txtdeveloper.Text = UCase(Datadeveloper.Recordset!kode_developer)
    zKDdeveloper = UCase(Datadeveloper.Recordset!kode_developer)
    DTPicker1.SetFocus
  Else
    FrmListdev.Show
  End If
End Sub

Private Sub txtdimb_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtimb.Text = txtimb.Text + (-txtdimb.Text)
If KeyCode = 13 Then txtdser.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtdjalan_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtjalan.Text = txtjalan.Text + (-txtdjalan.Text)
If KeyCode = 13 Then txtdair.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtdlistrik_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtlistrik.Text = txtlistrik.Text + (-txtdlistrik.Text)
If KeyCode = 13 Then txtdjalan.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtdser_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
If KeyCode = 13 Then txtser.Text = txtser.Text + (-txtdser.Text)
If KeyCode = 13 Then txtdlistrik.SetFocus
On Error Resume Next
On Error GoTo 0
ErrorHandler:
End Sub

Private Sub txtair_LostFocus()
If IsNumeric(txtair) Then txtair.Text = Format(txtair.Text, "###,###,##0")
End Sub

Private Sub txtbiayaapht_LostFocus()
If IsNumeric(txtbiayaapht) Then txtbiayaapht.Text = Format(txtbiayaapht.Text, "###,###,##0")
End Sub

Private Sub txtbiayan_LostFocus()
If IsNumeric(txtbiayan) Then txtbiayan.Text = Format(txtbiayan.Text, "###,###,##0")
End Sub

Private Sub txtimb_Change()
If txtimb.Text = "0" Then txtuimb.Text = "-1"
End Sub

Private Sub txtimb_LostFocus()
If IsNumeric(txtimb) Then txtimb.Text = Format(txtimb.Text, "###,###,##0")
End Sub

Private Sub txtjalan_Change()
If txtjalan.Text = "0" Then txtujalan.Text = "-1"
End Sub

Private Sub txtjalan_LostFocus()
If IsNumeric(txtjalan) Then txtjalan.Text = Format(txtjalan.Text, "###,###,##0")
End Sub

Private Sub txtlistrik_Change()
If txtlistrik.Text = "0" Then txtulistrik.Text = "-1"
End Sub

Private Sub txtlistrik_LostFocus()
If IsNumeric(txtlistrik) Then txtlistrik.Text = Format(txtlistrik.Text, "###,###,##0")
End Sub

Private Sub txtnotaris_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    If txtnotaris <> "" Then
      Call CariDatanot
    Else
      FrmListnot.Show
    End If
  End If
End Sub
Public Sub CariDatanot()
  Datanotaris.Recordset.FindFirst "kode_notaris = '" & txtnotaris.Text & "'"
  If Not Datanotaris.Recordset.NoMatch Then
    txtnotaris.Text = UCase(Datanotaris.Recordset!kode_notaris)
    zKDnotaris = UCase(Datanotaris.Recordset!kode_notaris)
  Else
    FrmListnot.Show
  End If
End Sub

Private Sub txtser_Change()
If txtser.Text = "0" Then TxtUser.Text = "-1"
End Sub

Private Sub txtser_LostFocus()
If IsNumeric(txtser) Then txtser.Text = Format(txtser.Text, "###,###,##0")
End Sub

Private Sub vbButton1_Click()
Cr.ReportFileName = App.Path & "\PENCAIRAN NOTARIS PERTANGGAL.rpt"
Cr.WindowTitle = "Dana Jaminan Pencairan Report"
Cr.WindowShowRefreshBtn = True
Cr.Destination = crptToWindow
Cr.Action = 1
End Sub

Private Sub vbButton2_Click()
cr2.ReportFileName = App.Path & "\PENCAIRAN DEVELOPER PERTANGGAL.rpt"
cr2.WindowTitle = "Dana Jaminan Pencairan Report"
cr2.WindowShowRefreshBtn = True
cr2.Destination = crptToWindow
cr2.Action = 1
End Sub

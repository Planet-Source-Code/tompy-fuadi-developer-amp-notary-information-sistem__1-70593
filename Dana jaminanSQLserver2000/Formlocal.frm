VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Formlocal 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Menu Kredit"
   ClientHeight    =   8130
   ClientLeft      =   3315
   ClientTop       =   1605
   ClientWidth     =   8340
   Icon            =   "Formlocal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin DanaJaminan.vbButton cmdkeluar 
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   7560
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
      MICON           =   "Formlocal.frx":0582
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdbatal2 
      Height          =   375
      Left            =   6840
      TabIndex        =   36
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":059E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdhapus2 
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":05BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdubah2 
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":05D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdcari2 
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Top             =   6000
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":05F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdsimpan2 
      Height          =   375
      Left            =   6840
      TabIndex        =   19
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":060E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdtambah2 
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":062A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DanaJaminan.vbButton cmdbatal 
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":0646
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
      Left            =   6840
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":0662
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
      Left            =   6840
      TabIndex        =   35
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":067E
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
      Left            =   6840
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":069A
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
      Left            =   6840
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":06B6
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
      Left            =   6840
      TabIndex        =   0
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "Formlocal.frx":06D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtno2 
      Height          =   285
      Left            =   2520
      TabIndex        =   34
      Text            =   "00007.01.01."
      Top             =   5400
      Width           =   3855
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
      TabIndex        =   33
      Top             =   5760
      Width           =   3855
   End
   Begin VB.TextBox txtbiayan 
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   6840
      Width           =   3855
   End
   Begin VB.TextBox txtbiayaapht 
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   7200
      Width           =   3855
   End
   Begin VB.TextBox txtstatus2 
      Height          =   285
      Left            =   2040
      TabIndex        =   30
      Text            =   "K"
      Top             =   7200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtstatus 
      Height          =   285
      Left            =   2040
      TabIndex        =   26
      Text            =   "K"
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cbjenis 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtnama 
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtair 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtjalan 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txtlistrik 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox txtser 
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
      TabIndex        =   9
      Top             =   3240
      Width           =   3855
   End
   Begin VB.TextBox txtimb 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
   End
   Begin VB.TextBox txtmak 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtno 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "00007.01.01."
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "INPUT DANA JAMINAN NOTARIS :"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   37
      Top             =   4680
      Width           =   8055
      Begin VB.TextBox Txtnamanot 
         Height          =   285
         Left            =   3840
         TabIndex        =   69
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtnotaris 
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtuapht 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   66
         Text            =   "1"
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox txtunotaris 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   65
         Text            =   "1"
         Top             =   2160
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   255
         Left            =   2400
         TabIndex        =   59
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   30196
      End
      Begin VB.TextBox txtnomer2 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtstatusapht 
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Text            =   "K"
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
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
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C00000&
         Caption         =   "&Notaris     "
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
         Top             =   1800
         Width           =   735
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
         TabIndex        =   42
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   41
         Top             =   2160
         Width           =   1335
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
         TabIndex        =   40
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C00000&
         Caption         =   "&Nama Debitur  "
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
         TabIndex        =   39
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Debitur  "
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
         TabIndex        =   38
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "INPUT DANA JAMINAN DEVELOPER :"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   8055
      Begin VB.TextBox Txtnamadev 
         Height          =   285
         Left            =   3720
         TabIndex        =   68
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtdeveloper 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtuair 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   64
         Text            =   "1"
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox txtujalan 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   63
         Text            =   "1"
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox txtulistrik 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   62
         Text            =   "1"
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   61
         Text            =   "1"
         Top             =   3240
         Width           =   255
      End
      Begin VB.TextBox txtuimb 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6360
         TabIndex        =   60
         Text            =   "1"
         Top             =   2880
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   30196
      End
      Begin VB.TextBox txtnomer 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtstatusair 
         Height          =   285
         Left            =   1920
         TabIndex        =   56
         Text            =   "K"
         Top             =   4320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatusjalan 
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Text            =   "K"
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatuslistrik 
         Height          =   285
         Left            =   1920
         TabIndex        =   28
         Text            =   "K"
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtstatusser 
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Text            =   "K"
         Top             =   3240
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
         TabIndex        =   67
         Top             =   240
         Width           =   1275
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
         TabIndex        =   58
         Top             =   360
         Width           =   1455
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
         TabIndex        =   55
         Top             =   4320
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
         TabIndex        =   54
         Top             =   3960
         Width           =   1935
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
         TabIndex        =   53
         Top             =   3600
         Width           =   1815
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
         TabIndex        =   52
         Top             =   3240
         Width           =   1695
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
         TabIndex        =   51
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C00000&
         Caption         =   "&Maksimal Kredit "
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
         Top             =   2520
         Width           =   1695
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
         TabIndex        =   49
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C00000&
         Caption         =   "&Developer   "
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
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C00000&
         Caption         =   "&Jenis Kredit "
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
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C00000&
         Caption         =   "&Nama Debitur  "
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
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C00000&
         Caption         =   "&Nomer Debitur  "
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
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Formlocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nomer As Integer
Dim zTgl As Date
Dim RSkredit As New ADODB.Recordset
Dim RSnomer As New ADODB.Recordset
Dim RSdeveloper As New ADODB.Recordset

Private Sub CmdBatal_Click()
Call PosisiAwal
End Sub

Private Sub cmdbatal2_Click()
Call PosisiAwal2
End Sub

Private Sub CmdCari_Click()
Call CariDatabase
Call TxtHidup
cmdubah.Enabled = True
cmdhapus.Enabled = True
End Sub

Public Sub CariDatabase()
    MDIForm1.Enabled = False
    FrmListBase.Show
End Sub

Private Sub cmdcari2_Click()
Call CariDatabase2
Call TxtHidup2
cmdubah.Enabled = True
cmdhapus.Enabled = True
End Sub

Public Sub CariDatabase2()
 MDIForm1.Enabled = False
 FrmListBaseOverseas.Show
End Sub

Private Sub cmdhapus_Click()
SQL = "select * from datadeveloper where nomer='" & Trim(txtnomer.Text) & "'"
If RSdeveloper.State > 0 Then RSdeveloper.Close
RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Apa Anda Yakin Mau Hapus Data Ini?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RSdeveloper.Delete
    MsgBox "Data Sudah Dihapus.", vbInformation, "Informasi"
    RSdeveloper.Close
End If
TxtKosong
cmdhapus.Enabled = False
End Sub
Private Sub PosisiAwal()
  Call TxtKosong
  Call TxtMati
  cmdsimpan.Enabled = False
  cmdhapus.Enabled = False
  CmdCari.Enabled = True
  cmdtambah.Enabled = True
  cmdubah.Enabled = True
End Sub

Private Sub PosisiAwal2()
  Call TxtKosong2
  Call TxtMati2
  cmdsimpan2.Enabled = False
  cmdhapus2.Enabled = False
  cmdcari2.Enabled = True
  cmdtambah2.Enabled = True
  cmdubah2.Enabled = True
End Sub

Private Sub cmdhapus2_Click()
SQL = "select * from datanotaris where nomer='" & Trim(txtnomer2.Text) & "'"
If RSdeveloper.State > 0 Then RSdeveloper.Close
RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
Tanya = MsgBox("Apa Anda Yakin Mau Hapus Data Ini?", vbQuestion + vbYesNo, "Konfirmasi Hapus")
If Tanya = vbYes Then
    RSdeveloper.Delete
    MsgBox "Data Sudah Dihapus.", vbInformation, "Informasi"
    RSdeveloper.Close
End If
TxtKosong2
cmdhapus2.Enabled = False
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
    .Fields("debetjaminanimb") = Trim(txtimb.Text)
    .Fields("statusimb") = Trim(txtstatus.Text)
    .Fields("unitimb") = Trim(txtuimb.Text)
    .Fields("jaminansertifikat") = Trim(txtser.Text)
    .Fields("debetjaminansertifikat") = Trim(txtser.Text)
    .Fields("statussertifikat") = Trim(txtstatusser.Text)
    .Fields("unitsertifikat") = Trim(TxtUser.Text)
    .Fields("jaminanlistrik") = Trim(txtlistrik.Text)
    .Fields("debetjaminanlistrik") = Trim(txtlistrik.Text)
    .Fields("statuslistrik") = Trim(txtstatuslistrik.Text)
    .Fields("unitlistrik") = Trim(txtulistrik.Text)
    .Fields("jaminanbestekjalan") = Trim(txtjalan.Text)
    .Fields("debetjaminanbestekjalan") = Trim(txtjalan.Text)
    .Fields("statusjalan") = Trim(txtstatusjalan.Text)
    .Fields("unitjalan") = Trim(txtujalan.Text)
    .Fields("jaminanbestekair") = Trim(txtair.Text)
    .Fields("debetjaminanbestekair") = Trim(txtair.Text)
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
    .Fields("debetbiayanotaris") = Trim(txtbiayan.Text)
    .Fields("statusnotaris") = Trim(txtstatus2.Text)
    .Fields("unitnotaris") = Trim(txtunotaris.Text)
    .Fields("biayaapht") = Trim(txtbiayaapht.Text)
    .Fields("debetbiayaapht") = Trim(txtbiayaapht.Text)
    .Fields("statusapht") = Trim(txtstatusapht.Text)
    .Fields("unitapht") = Trim(txtuapht.Text)
End With
End Sub

Private Sub cmdtambah_Click()
  Call TxtHidup
  cmdtambah.Enabled = False
  cmdubah.Enabled = False
  cmdsimpan.Enabled = True
  RSdeveloper.Open "datadeveloper", cn, adOpenDynamic, adLockOptimistic
  RSdeveloper.MoveLast
  txtnomer = "K" + Format(Val(Right(RSdeveloper!Nomer, 8)) + 1, "0000000#")
  txtno.SetFocus
  RSdeveloper.Close
End Sub
Private Sub TxtKosong()
     txtnomer.Text = ""
     txtno.Text = "00007.01.01."
     txtnama.Text = ""
     cbjenis.ListIndex = -1
     txtdeveloper.Text = ""
     txtmak.Text = ""
     txtimb.Text = ""
     txtser.Text = ""
     txtlistrik.Text = ""
     txtjalan.Text = ""
     txtair.Text = ""
End Sub
Private Sub TxtMati()
  txtno.Enabled = False
  txtnama.Enabled = False
  cbjenis.Enabled = False
  txtdeveloper.Enabled = False
  txtmak.Enabled = False
  txtimb.Enabled = False
  txtser.Enabled = False
  txtlistrik.Enabled = False
  txtjalan.Enabled = False
  txtair.Enabled = False
  txtstatus.Enabled = False
  txtstatusser.Enabled = False
  txtstatuslistrik.Enabled = False
  txtstatusjalan.Enabled = False
  txtstatusair.Enabled = False
End Sub
Private Sub TxtHidup()
  txtnomer.Enabled = True
  txtno.Enabled = True
  txtnama.Enabled = True
  cbjenis.Enabled = True
  txtdeveloper.Enabled = True
  txtmak.Enabled = True
  txtimb.Enabled = True
  txtser.Enabled = True
  txtlistrik.Enabled = True
  txtjalan.Enabled = True
  txtair.Enabled = True
  txtstatus.Enabled = True
  txtstatusser.Enabled = True
  txtstatuslistrik.Enabled = True
  txtstatusjalan.Enabled = True
  txtstatusair.Enabled = True
End Sub

Private Sub cmdtambah2_Click()
  Call TxtHidup2
  cmdtambah2.Enabled = False
  cmdubah2.Enabled = False
  cmdsimpan2.Enabled = True
End Sub

Private Sub cmdubah_Click()
  'Edit data
    SQL = "select * from datadeveloper where nomer='" & Trim(txtnomer.Text) & "'"
    If RSdeveloper.State > 0 Then RSdeveloper.Close
    RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
    With RSdeveloper
        Simpan
        .Update
        .Close
    End With
    MsgBox "Data Sudah Diubah.", vbInformation, "Information"
    TxtKosong
    cmdubah.Enabled = False
End Sub

Private Sub cmdubah2_Click()
'Edit data
    SQL = "select * from datanotaris where nomer='" & Trim(txtnomer2.Text) & "'"
    If RSdeveloper.State > 0 Then RSdeveloper.Close
    RSdeveloper.Open SQL, cn, adOpenDynamic, adLockOptimistic
    With RSdeveloper
        Simpan2
        .Update
        .Close
    End With
    MsgBox "Data Sudah Diubah.", vbInformation, "Information"
    TxtKosong2
    cmdubah2.Enabled = False
End Sub

Private Sub Form_Load()
  Me.Width = 8430
  Me.Height = 8535
  Me.Top = 1245
  Me.Left = 3270
DTPicker1 = Date
DTPicker2 = Date
lbtgl.Caption = Format(Date, "dd/mm/yyyy")
  zTgl = Format(Date, "dd/mm/yyyy")
Bukakoneksi
cn.CursorLocation = adUseClient
SQL = "Select * from tabelkredit order by jeniskredit"
Set RSkredit = cn.Execute(SQL)
Do While Not RSkredit.EOF
    cbjenis.AddItem RSkredit!jeniskredit
    RSkredit.MoveNext
Loop
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
MDIForm1.Enabled = False
FrmListdev.Show
End Sub

Private Sub txtimb_LostFocus()
If IsNumeric(txtimb) Then txtimb.Text = Format(txtimb.Text, "###,###,##0")
End Sub

Private Sub txtjalan_LostFocus()
If IsNumeric(txtjalan) Then txtjalan.Text = Format(txtjalan.Text, "###,###,##0")
End Sub

Private Sub txtlistrik_LostFocus()
If IsNumeric(txtlistrik) Then txtlistrik.Text = Format(txtlistrik.Text, "###,###,##0")
End Sub

Private Sub txtmak_LostFocus()
  If IsNumeric(txtmak) Then txtmak.Text = Format(txtmak.Text, "###,###,##0")
End Sub

Private Sub txtno_LostFocus()
txtno2.Text = txtno.Text
End Sub

Private Sub txtnama_LostFocus()
txtnama2.Text = txtnama.Text
End Sub

Private Sub txtnomer_lostfocus()
txtnomer2.Text = txtnomer.Text
End Sub

Private Sub DTPicker1_LostFocus()
DTPicker2.Value = DTPicker1.Value
End Sub

Private Sub TxtKosong2()
     txtnomer2.Text = ""
     txtno2.Text = "00007.01.01."
     txtnama2.Text = ""
     txtnotaris.Text = ""
     txtbiayan.Text = ""
     txtbiayaapht.Text = ""
End Sub
Private Sub TxtMati2()
  txtnomer.Enabled = False
  txtno2.Enabled = False
  txtnama2.Enabled = False
  txtnotaris.Enabled = False
  txtbiayan.Enabled = False
  txtbiayaapht.Enabled = False
  txtstatus2.Enabled = False
  txtstatusapht.Enabled = False
End Sub
Private Sub TxtHidup2()
  txtnomer.Enabled = True
  txtno2.Enabled = True
  txtnama2.Enabled = True
  txtnotaris.Enabled = True
  txtbiayan.Enabled = True
  txtbiayaapht.Enabled = True
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
 MDIForm1.Enabled = False
 FrmListnot.Show
End Sub

Private Sub txtser_LostFocus()
If IsNumeric(txtser) Then txtser.Text = Format(txtser.Text, "###,###,##0")
End Sub

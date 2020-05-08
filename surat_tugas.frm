VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form surat_tugas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SURAT LEMBUR TDC KISEL SRAGEN"
   ClientHeight    =   10110
   ClientLeft      =   8580
   ClientTop       =   1620
   ClientWidth     =   20370
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   PaletteMode     =   2  'Custom
   Picture         =   "surat_tugas.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cari"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cari"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6960
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19800
      Top             =   0
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   8280
      TabIndex        =   44
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Format          =   94109696
      CurrentDate     =   43466
      MinDate         =   36526
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Menu Utama"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8055
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   12135
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FFFF&
         Cancel          =   -1  'True
         Caption         =   "CETAK DATA SURAT TUGAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   7320
         UseMaskColor    =   -1  'True
         Width           =   5535
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   6360
         TabIndex        =   32
         Top             =   840
         Width           =   5655
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   48
            Top             =   4560
            Width           =   3615
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   42
            Top             =   2760
            Width           =   3615
         End
         Begin VB.ComboBox combo5 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            TabIndex        =   40
            Top             =   2160
            Width           =   3615
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   38
            Top             =   1560
            Width           =   3615
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   34
            Text            =   "Jl. Raya Sragen Km. 5 Solo"
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   33
            Text            =   "TDC KISEL SRAGEN"
            Top             =   360
            Width           =   3615
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   1800
            TabIndex        =   46
            Top             =   3960
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            _Version        =   393216
            Format          =   93913088
            CurrentDate     =   43466
            MinDate         =   33239
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000005&
            Caption         =   "Tempat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   4560
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Caption         =   "Sampai Tanggal Tanggal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000005&
            Caption         =   "Mulai Tanggal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label15 
            BackColor       =   &H80000005&
            Caption         =   "Keperluan Lain*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "Keperluan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "No.Telfon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   39
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Unit Kerja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Caption         =   "DATA PENUGASAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   24
         Top             =   5040
         Width           =   6135
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   28
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   27
            Top             =   960
            Width           =   3615
         End
         Begin VB.ComboBox Combo3 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   26
            Top             =   1560
            Width           =   3615
         End
         Begin VB.ComboBox Combo4 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   25
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Departemen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   37
            Top             =   2160
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000005&
            Caption         =   "Nama yang Ditugasi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "NIP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   30
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "DATA PEMBERI TUGAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   6135
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   50
            Top             =   1560
            Width           =   3615
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   23
            Text            =   "TDC KISEL SRAGEN"
            Top             =   3480
            Width           =   3615
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   20
            Top             =   2280
            Width           =   3615
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc2"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   18
            Top             =   2880
            Width           =   3615
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   16
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   15
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000005&
            Caption         =   "No Surat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Unit Kerja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Departemen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Jabatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "NIP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000005&
            Caption         =   "Nama Pemberi Tugas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6120
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6720
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   6120
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FFFF&
         Caption         =   "UBAH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6720
         Width           =   2655
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "FORM PENGISIAN SURAT TUGAS"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   0
         Width           =   9975
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000C0&
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   12135
      End
      Begin VB.Label Label5 
         Caption         =   "Kewarganegaraa /Agama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4440
         Width           =   2655
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94044160
      CurrentDate     =   43194
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   12240
      Top             =   8880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "D:\KISEL\lap_surat.rpt"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94044160
      CurrentDate     =   43140
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   12120
      TabIndex        =   0
      Top             =   1560
      Width           =   8055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "surat_tugas.frx":3F95
         Height          =   5775
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
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
            Name            =   "Tahoma"
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
               LCID            =   1057
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
               LCID            =   1057
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
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         Caption         =   "CARI DATA BERDASARKAN NAMA/NOMOR SURAT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   7575
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000C0&
         Caption         =   "Label6"
         Height          =   855
         Index           =   0
         Left            =   -4440
         TabIndex        =   52
         Top             =   -240
         Width           =   12495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\KISEL\Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\KISEL\Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "datamas"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\KISEL\Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\KISEL\Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "combo"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "SURAT TUGAS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   60
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000C0&
      Height          =   1215
      Index           =   3
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   20775
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   54
      Top             =   9720
      Width           =   19815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9480
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "surat_tugas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================FORM SURAT CODE ===========================
     '======================= LUBIS TEGUH ===========================

'MENDEKLARASIKAN DATACOMBO
Dim DataCombo As New ADODB.Recordset

'SCRIPT AUTONUMBER SURAT

'MENAMPILKAN DATA PADA DATABASE KE COMBO
Sub tambahcom()
Adodc2.ConnectionString = conn.ConnectionString
Adodc2.RecordSource = "select* from combo"

For Each gosong In Me.Controls
If TypeOf gosong Is ComboBox Then
gosong.Text = ""
With Adodc2.Recordset
    Do While Not .EOF
    On Error Resume Next
    Combo1.AddItem !jabatan
    Combo2.AddItem !departemen
    Combo3.AddItem !jabatan
    Combo4.AddItem !departemen
    Combo5.AddItem !keperluan
    .MoveNext
    Loop
End With
End If
Next
End Sub

'CLEAR FORM
Sub bersih()
Text1 = ""
Text2 = ""
Text3 = ""
'Text4 = ""
Text5 = ""
Text6 = ""
'Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text11 = ""
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo5 = ""
Combo5 = ""
'Option1.Value = False
'Option2.Value = False
DTPicker3.Value = Now
DTPicker4.Value = Now
End Sub

'ENABLE TRUE FORM
Sub tambah()
Text1.Enabled = True
'Text2.Enabled = True
'Text3.Enabled = True
'Text5.Enabled = True
'Text6.Enabled = True
'Text8.Enabled = True
'Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
'Combo1.Enabled = True
'Combo2.Enabled = True
'Combo3.Enabled = True
'Combo4.Enabled = True
Combo5.Enabled = True
DTPicker3.Enabled = True
DTPicker4.Enabled = True
Command8.Enabled = True
Command6.Enabled = True
Text1.SetFocus
End Sub

'FORMAT WAKTU DATAGRID
Sub pormat()
'DataGrid1.Columns(4).NumberFormat = ("DD/MM/YYYY")
'DataGrid1.Columns(14).NumberFormat = ("DD/MM/YYYY")
End Sub


'JIKA TOMBOL TAMBAH DI KLIK PANGGIL FUNGSI TERSEBUT
Private Sub Command1_Click()
Call tambahcom
Call tambah
Call bersih
'cari_penduduk.Show
'Call AutoNumber
Command2.Enabled = True
End Sub

'SIMPAN DATA
Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text11 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!nomor = Text1.Text
Adodc1.Recordset!nama_pt = Text2.Text
Adodc1.Recordset!nip = Text3.Text
Adodc1.Recordset!jabatan = Combo1.Text
Adodc1.Recordset!departemen = Combo2.Text
Adodc1.Recordset!unit_kerja = Text4.Text
Adodc1.Recordset!nama_t = Text5.Text
Adodc1.Recordset!nip_t = Text6.Text
Adodc1.Recordset!jabatan_t = Combo3.Text
Adodc1.Recordset!departemen_t = Combo4.Text
Adodc1.Recordset!unit_kerja_t = Text7.Text
Adodc1.Recordset!alamat = Text8.Text
Adodc1.Recordset!no_tlfn = Text9.Text
Adodc1.Recordset!keperluan = Combo5.Text
Adodc1.Recordset!keperluan_lain = Text10.Text
Adodc1.Recordset!tanggal_mulai = DTPicker3.Value
Adodc1.Recordset!tanggal_selesai = DTPicker4.Value
Adodc1.Recordset!tempat = Text11.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL DISIMPAN !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
Call pormat
End Sub

'UBAH DATA
Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text11 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA UBAH !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset!nomor = Text1.Text
Adodc1.Recordset!nama_pt = Text2.Text
Adodc1.Recordset!nip = Text3.Text
Adodc1.Recordset!jabatan = Combo1.Text
Adodc1.Recordset!departemen = Combo2.Text
Adodc1.Recordset!unit_kerja = Text4.Text
Adodc1.Recordset!nama_t = Text5.Text
Adodc1.Recordset!nip_t = Text6.Text
Adodc1.Recordset!jabatan_t = Combo3.Text
Adodc1.Recordset!departemen_t = Combo4.Text
Adodc1.Recordset!unit_kerja_t = Text7.Text
Adodc1.Recordset!alamat = Text8.Text
Adodc1.Recordset!no_tlfn = Text9.Text
Adodc1.Recordset!keperluan = Combo5.Text
Adodc1.Recordset!keperluan_lain = Text10.Text
Adodc1.Recordset!tanggal_mulai = DTPicker3.Value
Adodc1.Recordset!tanggal_selesai = DTPicker4.Value
Adodc1.Recordset!tempat = Text11.Text
Adodc1.Recordset.Update
Command2.Enabled = True
Call bersih
MsgBox "DATA ANDA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'HAPUS DATA
Private Sub Command4_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("Apakah Anda yakin akan menghapus data?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call bersih
MsgBox "DATA ANDA BERHASIL DIHAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
            End If
End If
End Sub

'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama_pt like '%" + Me.Text14.Text + "%' or nomor like '%" + Me.Text14.Text + "%' or nama_t like '%" + Me.Text14.Text + "%'"
End Sub

Private Sub Command6_Click()
Call tambah
cari_karyawan.Show
End Sub

Private Sub Command8_Click()
cari_karyawan_2.Show
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text14_Change()
If Text14.Text = "" Then
Adodc1.Refresh
Else
'wkwk
End If
End Sub

'PINDAH DATA DARI DATAGRIDVIEW KE TEXTBOX
Private Sub DataGrid1_Click()
 Text1.Text = Adodc1.Recordset!nomor
 Text2.Text = Adodc1.Recordset!nama_pt
Text3.Text = Adodc1.Recordset!nip
Combo1.Text = Adodc1.Recordset!jabatan
Combo2.Text = Adodc1.Recordset!departemen
Text4.Text = Adodc1.Recordset!unit_kerja
Text5.Text = Adodc1.Recordset!nama_t
Text6.Text = Adodc1.Recordset!nip_t
Combo3.Text = Adodc1.Recordset!jabatan_t
Combo4.Text = Adodc1.Recordset!departemen_t
Text7.Text = Adodc1.Recordset!unit_kerja_t
Text8.Text = Adodc1.Recordset!alamat
Text9.Text = Adodc1.Recordset!no_tlfn
Combo5.Text = Adodc1.Recordset!keperluan
Text10.Text = Adodc1.Recordset!keperluan_lain
DTPicker3.Value = Adodc1.Recordset!tanggal_mulai
DTPicker4.Value = Adodc1.Recordset!tanggal_selesai
Text11.Text = Adodc1.Recordset!tempat
Command2.Enabled = False
End Sub

'MUNCULKAN LAPORAN
Private Sub Command7_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA CETAK !", vbInformation, "PERHATIAN !"
Else
With cr1
    .SelectionFormula = "{datamas.nomor}='" & Text1.Text & "'"
    .ReportFileName = App.Path & "\surat_tugas.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End If
'cr1.Connect = "dsn=database"
'cr1.ReportFileName = App.Path & "\lap_surat.rpt"
'cr1.Action = 1
End Sub

Private Sub DATA_ADMIN_Click()
admin.Show
End Sub

'FUNGSI AKTIF OTOMATIS SAAT FORM DIBUKA
Private Sub Form_Load()
'combo 1
Call tambahcom
Call bersih
Call pormat
End Sub


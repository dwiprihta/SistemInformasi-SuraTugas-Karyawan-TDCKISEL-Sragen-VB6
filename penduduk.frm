VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form karyawan 
   BackColor       =   &H80000005&
   Caption         =   "FORM INPUT DATA PENDUDUK"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   ScaleHeight     =   10755
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   375
      Left            =   9240
      TabIndex        =   44
      Top             =   3720
      Width           =   3615
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   19560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "CETAK"
      Height          =   495
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6240
      Width           =   1455
   End
   Begin VB.ComboBox Combo7 
      DataSource      =   "Adodc2"
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
      Left            =   8400
      TabIndex        =   29
      Text            =   "Combo7"
      Top             =   6360
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19920
      Top             =   0
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   18360
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   17160
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      RecordSource    =   "karyawan"
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Textcari 
      Appearance      =   0  'Flat
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
      Left            =   14160
      TabIndex        =   25
      Top             =   6240
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "penduduk.frx":0000
      Height          =   3255
      Left            =   480
      TabIndex        =   23
      Top             =   6960
      Width           =   19335
      _ExtentX        =   34105
      _ExtentY        =   5741
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
      Height          =   8895
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   20415
      Begin VB.ComboBox Combo8 
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
         Left            =   16200
         TabIndex        =   46
         Top             =   1800
         Width           =   3615
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   16200
         TabIndex        =   45
         Top             =   1200
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   16200
         TabIndex        =   43
         Top             =   3720
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   131072001
         CurrentDate     =   43501
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   42
         Top             =   4320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   131072001
         CurrentDate     =   43501
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2520
         TabIndex        =   41
         Top             =   3000
         Width           =   3615
      End
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
         Left            =   9240
         TabIndex        =   39
         Top             =   3600
         Width           =   3615
      End
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
         Left            =   2520
         TabIndex        =   37
         Top             =   3720
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
         Height          =   375
         Left            =   16200
         TabIndex        =   33
         Text            =   "TDC KISEL SRAGEN"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox combo6 
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
         Left            =   16200
         TabIndex        =   32
         Top             =   3120
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         Caption         =   "CETAK DATA BERDASARAN"
         Height          =   975
         Left            =   8280
         TabIndex        =   28
         Top             =   5160
         Width           =   5535
      End
      Begin VB.Frame Frame4 
         Caption         =   "CARI DATA BERDASARKAN NAMA/NOMOR KTP"
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
         Left            =   14040
         TabIndex        =   24
         Top             =   5160
         Width           =   5775
      End
      Begin VB.CommandButton Command7 
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5520
         Width           =   1575
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   9240
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "TAMBAH DATA"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5520
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "Perempuan"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Laki-Laki"
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
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1800
         Width           =   3615
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
         Left            =   2520
         TabIndex        =   4
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   9240
         TabIndex        =   3
         Top             =   1080
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
         Height          =   375
         Left            =   9240
         TabIndex        =   2
         Top             =   2400
         Width           =   3615
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   9240
         TabIndex        =   1
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "FORM INPUT DATA KARYAWAN"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000C0&
         Height          =   735
         Index           =   0
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   20415
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Caption         =   "Telpon"
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
         Left            =   6960
         TabIndex        =   40
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "Jumlah Anak"
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
         Left            =   6960
         TabIndex        =   38
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Masuk"
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
         Left            =   13920
         TabIndex        =   36
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "Status Karyawan"
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
         Left            =   13920
         TabIndex        =   35
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   13920
         TabIndex        =   34
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   13440
         X2              =   13440
         Y1              =   960
         Y2              =   4680
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Agama"
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
         Left            =   480
         TabIndex        =   31
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   3
         Left            =   13920
         TabIndex        =   27
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   6600
         X2              =   6600
         Y1              =   960
         Y2              =   4680
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Pendidikan Terakhir"
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
         Index           =   2
         Left            =   6960
         TabIndex        =   20
         Top             =   1080
         Width           =   1815
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
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Nama"
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
         Index           =   1
         Left            =   480
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Jenis Kelamin"
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
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Tempat Lahir"
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
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   6960
         TabIndex        =   15
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Kewarganegaraan"
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
         Left            =   6960
         TabIndex        =   14
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   13920
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Lahir"
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
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Status Pernikahan"
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
         Left            =   6960
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
   End
End
Attribute VB_Name = "karyawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'======================== FORM PENDUDUK CODE ==========================='
     '======================= LUBIS PAMBUDI ==========================='
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
    Combo1.AddItem !agama
    Combo2.AddItem !pendidikan
    Combo3.AddItem !status_perkawinan
    Combo4.AddItem !departemen
    Combo5.AddItem !kewarganegaraan
    combo6.AddItem !status_krywn
    Combo8.AddItem !jabatan
    
    'Text9.Text = !lurah
    'Text10.Text = !camat
    Text7.AddItem !keperluan
    .MoveNext
    Loop
End With
End If
Next
End Sub

Private Sub Command3_Click()
 If Combo7 = "" Then
  MsgBox "SORTIR DAHULU DATA YANG AKAN ANDA TAMPILKAN !", vbInformation, "PERHATIAN !"
  ElseIf Combo7.Text = "SEMUA" Then
    CR1.Reset
With CR1
    .ReportFileName = App.Path & "\Data_penduduk.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
Else
'WKWK
With CR1
    .SelectionFormula = "{penduduk.no_kk}='" & Text1.Text & "'"
    .ReportFileName = App.Path & "\keluarga.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End If
End Sub

'FORMAT WAKTU DATAGRID
Sub pormat()
DataGrid1.Columns(4).NumberFormat = ("DD/MM/YYYY")
End Sub
     
'CLEAR FORM
Sub bersih()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text6 = ""
Text7 = ""
'Text8 = ""
Text9 = ""
Option1.Value = False
Option2.Value = False
DTPicker1.Value = Now
DTPicker2.Value = Now
Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo5 = ""
combo6 = ""
End Sub

'ENABLE TRUE FORM
Sub tambah()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Combo4.Enabled = True
Combo8.Enabled = True
'Text8.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Text9.Enabled = True
Combo5.Enabled = True
combo6.Enabled = True
End Sub

'TAMBAH
Private Sub Command1_Click()
Call bersih
Call tambah
End Sub

'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text5.Text + "%' or no_kk like '%" + Me.Text5.Text + "%'or nik like '%" + Me.Text5.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text5_Change()
If Text5.Text = "" Then
Adodc1.Refresh
Else
'wkwk
End If
End Sub

'PINDAH DATA DARI DATAGRIDVIEW KE TEXTBOX
Private Sub DataGrid1_Click()
Text1 = Adodc1.Recordset!no_kk
Text2 = Adodc1.Recordset!nik
Text3 = Adodc1.Recordset!nama
Text4 = Adodc1.Recordset!tempat_lahir
DTPicker2 = Adodc1.Recordset!tanggal_lahir
If Adodc1.Recordset!jk = "Laki-Laki" Then
    Option1.Value = True
ElseIf Adodc1.Recordset!jk = "Perempuan" Then
    Option2.Value = True
End If
Combo1 = Adodc1.Recordset!pekerjaan
Combo2 = Adodc1.Recordset!pendidikan
Combo3 = Adodc1.Recordset!alamat
Combo4 = Adodc1.Recordset!kewarganegaraan
combo_agama = Adodc1.Recordset!agama
Combo5 = Adodc1.Recordset!status_keluarga
combo6 = Adodc1.Recordset!status_perkawinan
Text6 = Adodc1.Recordset!no_paspor
Text7 = Adodc1.Recordset!no_kitas
Text8 = Adodc1.Recordset!nama_ayah
Text9 = Adodc1.Recordset!nama_ibu
End Sub

'LOAD
Private Sub Form_Load()
Call bersih
Call tambahcom

'NO TIME
Call pormat

'SORTIR
Combo7.AddItem ("SEMUA")
Combo7.AddItem ("KELUARGA")
End Sub

'SIMPAN DATA
Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!no_kk = Text1.Text
Adodc1.Recordset!nik = Text2.Text
Adodc1.Recordset!nama = Text3.Text
If Option1.Value = True Then
    Adodc1.Recordset!jk = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jk = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text4.Text
Adodc1.Recordset!tanggal_lahir = DTPicker2.Value
Adodc1.Recordset!pekerjaan = Combo1.Text
Adodc1.Recordset!pendidikan = Combo2.Text
Adodc1.Recordset!alamat = Combo3.Text
Adodc1.Recordset!kewarganegaraan = Combo4.Text
Adodc1.Recordset!agama = combo_agama.Text
Adodc1.Recordset!status_keluarga = Combo5.Text
Adodc1.Recordset!status_perkawinan = combo6.Text
Adodc1.Recordset!no_paspor = Text6.Text
Adodc1.Recordset!no_kitas = Text7.Text
Adodc1.Recordset!nama_ayah = Text8.Text
Adodc1.Recordset!nama_ibu = Text9.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL KAMI SIMPAN !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'UBAH
Private Sub Command6_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Or combo6 = "" Or combo_agama = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA UBAH !", vbInformation, "PERHATIAN !"
Else

Adodc1.Recordset!no_kk = Text1.Text
Adodc1.Recordset!nik = Text2.Text
Adodc1.Recordset!nama = Text3.Text
If Option1.Value = True Then
    Adodc1.Recordset!jk = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc1.Recordset!jk = Option2.Caption
End If
Adodc1.Recordset!tempat_lahir = Text4.Text
Adodc1.Recordset!tanggal_lahir = DTPicker2.Value
Adodc1.Recordset!pekerjaan = Combo1.Text
Adodc1.Recordset!pendidikan = Combo2.Text
Adodc1.Recordset!alamat = Combo3.Text
Adodc1.Recordset!kewarganegaraan = Combo4.Text
Adodc1.Recordset!agama = combo_agama.Text
Adodc1.Recordset!status_keluarga = Combo5.Text
Adodc1.Recordset!status_perkawinan = combo6.Text
Adodc1.Recordset!no_paspor = Text6.Text
Adodc1.Recordset!no_kitas = Text7.Text
Adodc1.Recordset!nama_ayah = Text8.Text
Adodc1.Recordset!nama_ibu = Text9.Text
Adodc1.Recordset.Update
Call bersih
MsgBox "DATA ANDA BERHASIL KAMI SIMPAN !", vbInformation, "INFORMASI !"
Adodc1.Refresh
End If
End Sub

'HAPUS DATA
Private Sub Command7_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or (Option1.Value = False And Option2.Value = False) Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Combo4 = "" Or Combo5 = "" Or combo6 = "" Or combo_agama = "" Or Text8 = "" Or Text9 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("Apakah Anda yakin akan menghapus data?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call bersih
MsgBox "DATA ANDA BERHASIL KAMI HAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
            End If
End If
End Sub



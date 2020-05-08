VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cari_karyawan_2 
   Caption         =   "CARI DATA KARYAWAN"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "UNTUK KEPERLUAN"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "SURAT LEMBUR"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SURAT TUGAS"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "CARI DATA BERDASARKAN NAMA"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   960
      Width           =   7575
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "cari_karyawan_2.frx":0000
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7223
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI DATA KARYAWAN PEMBERI TUGAS"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000C0&
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI DATA KARYAWAN"
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
      TabIndex        =   4
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "cari_karyawan_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CARI DATA
Private Sub Command5_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text14.Text + "%' or nik like '%" + Me.Text14.Text + "%'"
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
If Option1.Value = False And Option2.Value = False Then
MsgBox "PILIH DAHULU KEPERLUAN SURAT !", vbInformation, "PERHATIAN !"
End If

If Option1.Value = True Then
surat_tugas.Text2 = Adodc1.Recordset!nama
surat_tugas.Text3 = Adodc1.Recordset!nip
surat_tugas.Combo1 = Adodc1.Recordset!jabatan
surat_tugas.Combo2 = Adodc1.Recordset!departemen
surat_tugas.Show
Unload Me
End If

If Option2.Value = True Then
surat_lembur.Text2 = Adodc1.Recordset!nama
surat_lembur.Text3 = Adodc1.Recordset!nip
surat_lembur.Combo1 = Adodc1.Recordset!jabatan
surat_lembur.Combo2 = Adodc1.Recordset!departemen
surat_lembur.Show
Unload Me
End If
End Sub


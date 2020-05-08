VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form welcome 
   Caption         =   "SELAMAT DATANG (APLIKAIS PENGANTAR SURAT TDC KISEL SRAGEN)"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   Picture         =   "welcome.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   16680
      Top             =   0
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   16320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "SRAGEN"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9360
      TabIndex        =   2
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label88 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--/--/----"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   9015
   End
   Begin VB.Label Label77 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "--:--:--"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   11760
      TabIndex        =   0
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Menu master 
      Caption         =   "MASTER"
      Begin VB.Menu datpen 
         Caption         =   "DATA KARYAWAN"
      End
      Begin VB.Menu lemburrr 
         Caption         =   "LAPORAN SURAT LEMBUR"
      End
      Begin VB.Menu report 
         Caption         =   "LAPORAN SURAT TUGAS"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "TRANSAKSI SURAT"
      Begin VB.Menu surat 
         Caption         =   "SURAT TUGAS"
      End
      Begin VB.Menu lembur 
         Caption         =   "SURAT LEMBUR"
      End
   End
   Begin VB.Menu user2 
      Caption         =   "ADMINISTRATOR"
   End
   Begin VB.Menu out 
      Caption         =   "LOG-OUT"
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'======================= FORM HALAMAN DEPAN CODE===========================
     '======================= LUBIS TEGUH P ===========================
Private Sub Form_Load()
Label5 = login.Text1.Text
End Sub

Private Sub lembur_Click()
surat_lembur.Show
End Sub

Private Sub lemburrr_Click()
cr1.Reset
With cr1
    .ReportFileName = App.Path & "\data_surat2.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

Private Sub out_Click()
End
End Sub

'TAMPILKAN FORM TRANSAKSI
Private Sub surat_Click()
surat_tugas.Show
End Sub

Private Sub datpen_Click()
karyawan.Show
End Sub

'TAMPILKAN LAPORAN KESELURUHAN DATA
Private Sub report_Click()
cr1.Reset
With cr1
    .ReportFileName = App.Path & "\data_surat.rpt"
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub



'KELUAR APLIKASI
Private Sub keluar_Click()
xx = MsgBox("Apakah Anda yakin akan kelua dari aplikasi pengantar surat desa kroyo ?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
                    Unload Me
                Else
                    'NO NOTIF
            End If
End Sub

'TAMPILKAN PENGATURAN (NAMA LURAH & CAMAT)
Private Sub setting_Click()
'pengaturan.Show
End Sub

'SOURCE JAM BERJALAN
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
Label88.Caption = Format(Now, "dd MMMM yyyy")
End Sub

Private Sub user2_Click()
admin.Show
End Sub

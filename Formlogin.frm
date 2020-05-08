VERSION 5.00
Begin VB.Form login 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Formlogin.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "LOGIN"
      Height          =   615
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   360
      Picture         =   "Formlogin.frx":1872A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4740
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'======================= FORM LOGIN CODE===========================
     '======================= LUBIS TEGUH P ===========================
     
     
'JIKA TOMBOL LOGIN DI KLIK
Private Sub Command1_Click()

'panggil koneksi
Call Koneksi

'cek jika form masih kosong
If Text1.Text = "" Then
MsgBox "FORM USERNAME ANDA MASIH KOSONG !", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "FORM PASSWORD ANDA MASIH KOSONG !!!", vbCritical, "Perhatian"
Text2.SetFocus
Else

'cari data login di database admin
query = "select * from login where username='" & Text1.Text & "' and password='" & Text2.Text & "'"
RS.Open (query), conn
    If RS.EOF Then
    MsgBox "USERNAME ATAU PASSWORD ANDA SALAH !", vbExclamation, "Gagal !"
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Else
    
    'jika berhasil login masuk ke menu admin
    'Unload Me
    MsgBox "ANDA BERHASIL LOGIN !", vbInformation, "LOGIN SUKSES !"
    welcome.Show
    End If
End If
End Sub

'JIKA TOMBOL CANCEL DIKLIK
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Picture1_Click()

End Sub

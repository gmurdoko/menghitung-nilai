VERSION 5.00
Begin VB.Form frmmenu 
   Caption         =   "..:::::Menu"
   ClientHeight    =   6465
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   10695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dayu sanjaya -14753015-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Program Penilaian Mahasiswa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   5295
   End
   Begin VB.Menu mnudata 
      Caption         =   "Data"
      Begin VB.Menu mnumhs 
         Caption         =   "Data Mahasiswa"
      End
      Begin VB.Menu mnudosen 
         Caption         =   "Data Dosen"
      End
      Begin VB.Menu mnukuliah 
         Caption         =   "Data Matakuliah"
      End
   End
   Begin VB.Menu mnulaporan 
      Caption         =   "Laporan Nilai"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cn As New ADODB.Connection
Public Sub bukakoneksi()
strkoneksi = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & App.Path & "\data\NILAI.mdb;" & _
            "Persist Security Info=False "
cn.Open strkoneksi
If cn.State = 1 Then
    MsgBox " Program Dayu sanjaya ", vbInformation
Else
    MsgBox " Koneksi Gagal", vbInformation
End If
End Sub

Private Sub Form_Load()
bukakoneksi
End Sub
Private Sub mnudosen_Click()
frmdosen.Show 1
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnukuliah_Click()
frmmatkul.Show 1
End Sub

Private Sub mnulaporan_Click()
frmnilai.Show 1
End Sub

Private Sub mnumhs_Click()
frmmhs.Show 1
End Sub

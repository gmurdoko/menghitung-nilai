VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmhs 
   Caption         =   "..::::::Data Mahasiswa"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   14775
      Begin VB.TextBox txtnpm 
         Height          =   495
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtnama 
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   1080
         Width           =   4455
      End
      Begin VB.ComboBox cbojenis 
         Height          =   315
         ItemData        =   "frmmhs.frx":0000
         Left            =   2760
         List            =   "frmmhs.frx":000A
         TabIndex        =   12
         Text            =   "-------------------PILIH-------------------"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txttempat 
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtalamat 
         Height          =   735
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3960
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H8000000B&
         Height          =   1095
         Left            =   7680
         TabIndex        =   3
         Top             =   3840
         Width           =   6855
         Begin VB.CommandButton cmdtambah 
            Caption         =   "&Tambah"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdkeluar 
            Caption         =   "&Keluar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5520
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdhapus 
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2760
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdsimpan 
            Caption         =   "&Simpan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4080
            TabIndex        =   4
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdcari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   7560
         TabIndex        =   2
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5953
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker dtanggal 
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   3120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         Format          =   106299393
         CurrentDate     =   42528
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NPM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   19
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   1965
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   1845
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   16
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   15
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   7440
         X2              =   7440
         Y1              =   120
         Y2              =   6000
      End
   End
End
Attribute VB_Name = "frmmhs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcari_Click()
Dim rscari As New ADODB.Recordset
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from mahasiswa where npm='" & txtnpm.Text & "'", cn, adOpenStatic, adLockOptimistic
If rscari.RecordCount > 0 Then
    'MsgBox "Data Ada"
    txtnama = rscari.Fields("nama_mhs").Value
    cbojenis = rscari.Fields("jenis_kelamin").Value
    txttempat = rscari.Fields("tempat_lahir").Value
    dtanggal = rscari.Fields("tanggal_lahir").Value
    txtalamat = rscari.Fields("alamat").Value
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = False
Else
    'MsgBox "Data tidak Ada"
End If
End Sub

Private Sub Form_Load()
If cn.State = 0 Then bukakoneksi
isigrid
txtnama.Enabled = False
cbojenis.Enabled = False
txttempat.Enabled = False
dtanggal.Enabled = False
txtalamat.Enabled = False
    cmdedit.Enabled = False
    cmdhapus.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = True
End Sub
Public Sub isigrid()
Dim rsmahasiswa As New ADODB.Recordset
If rsmahasiswa.State = 1 Then rsmahasiswa.Close
rsmahasiswa.Open "select * from mahasiswa", cn, adOpenStatic, adLockOptimistic
If rsmahasiswa.RecordCount > 0 Then
Else

End If

With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .View = lvwReport
    .LabelEdit = lvwManual
    .ListItems.Clear
    .ColumnHeaders.Add 1, , "NPM"
    .ColumnHeaders.Add 2, , "Nama Mahasiswa"
    .ColumnHeaders.Add 3, , "Jenis Kelamin"
    .ColumnHeaders.Add 4, , "Tempat Lahir"
    .ColumnHeaders.Add 5, , "Tanggal Lahir"
    .ColumnHeaders.Add 6, , "Alamat"
    .ColumnHeaders(1).Width = 1200
    .ColumnHeaders(2).Width = 2500
    .ColumnHeaders(3).Width = 1200
    .ColumnHeaders(4).Width = 1200
    .ColumnHeaders(5).Width = 1200
    .ColumnHeaders(6).Width = 2000
  Do Until rsmahasiswa.EOF
    .ListItems.Add 1, , rsmahasiswa.Fields("NPM").Value & ""
    .ListItems(1).SubItems(1) = rsmahasiswa.Fields("nama_mhs").Value & ""
    .ListItems(1).SubItems(2) = rsmahasiswa.Fields("jenis_kelamin").Value & ""
    .ListItems(1).SubItems(3) = rsmahasiswa.Fields("tempat_lahir").Value & ""
    .ListItems(1).SubItems(4) = rsmahasiswa.Fields("tanggal_lahir").Value & ""
    .ListItems(1).SubItems(5) = rsmahasiswa.Fields("alamat").Value & ""
    rsmahasiswa.MoveNext
  Loop
End With
End Sub
Private Sub cmdedit_Click()
If cmdedit.Caption = "&Edit" Then
    cmdedit.Caption = "&Batal"
    cmdsimpan.Enabled = True
    cmdhapus.Enabled = False
    txtnpm.Enabled = True
txtnama.Enabled = True
   cbojenis.Enabled = True
    txttempat.Enabled = True
    dtanggal.Enabled = True
    txtalamat.Enabled = True
Else
   cmdedit.Caption = "&Edit"
   Form_Load
End If
End Sub
Private Sub cmdhapus_Click()
If MsgBox("Data mau dihapus ?", vbYesNo) = vbYes Then
cn.Execute "DELETE FROM mahasiswa WHERE npm='" & txtnpm & "'"
isigrid
End If
txtnpm.Text = ""
txtnama.Text = ""
cbojenis.Text = ""
dtanggal = ""
txttempat.Text = ""
txtalamat.Text = ""
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
If txtnpm.Text = "" Then
        MsgBox "npm belum diisi"
        txtnpm.SetFocus
ElseIf txtnama.Text = "" Then
        MsgBox "nama belum diisi"
        txtnama.SetFocus
Else
 
    If cmdedit.Caption = "&Batal" Then
        If MsgBox("Data mau diedit ?", vbYesNo) = vbYes Then
            cn.Execute "UPDATE mahasiswa SET nama_mhs='" & txtnama & "',jenis_kelamin='" & cbojenis _
            & "',tempat_lahir='" & txttempat & "',tanggal_lahir='" & dtanggal & "',alamat='" & txtalamat _
            & "' WHERE npm='" & txtnpm & "'"
            isigrid
            cmdedit_Click
        End If
    ElseIf cmdtambah.Caption = "&Batal" Then
        If MsgBox("Data mau ditambah ?", vbYesNo) = vbYes Then
            cn.Execute "INSERT INTO mahasiswa (NPM,nama_mhs,jenis_kelamin,tempat_lahir,tanggal_lahir,alamat) VALUES ('" & txtnpm _
            & "','" & txtnama & "','" & cbojenis & "','" & txttempat & "','" & dtanggal & "','" & txtalamat & "')"
            isigrid
            cmdtambah_Click
        End If
    End If
End If
End Sub

Private Sub cmdtambah_Click()
If cmdtambah.Caption = "&Tambah" Then
   cmdtambah.Caption = "&Batal"
   cmdsimpan.Enabled = True
txtnpm.Enabled = True
txtnama.Enabled = True
   cbojenis.Enabled = True
    txttempat.Enabled = True
    dtanggal.Enabled = True
    txtalamat.Enabled = True
Else
cmdtambah.Caption = "&Tambah"
Form_Load
End If
End Sub

Private Sub ListView1_DblClick()
 txtnpm.Text = ListView1.SelectedItem.Text
  txtnama.Text = ListView1.SelectedItem.SubItems(1)
  cbojenis.Text = ListView1.SelectedItem.SubItems(2)
  txttempat.Text = ListView1.SelectedItem.SubItems(3)
  dtanggal = ListView1.SelectedItem.SubItems(4)
  txtalamat.Text = ListView1.SelectedItem.SubItems(5)
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = False
End Sub

Private Sub txtnpm_LostFocus()
cmdcari_Click
End Sub




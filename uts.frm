VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmnilai 
   Caption         =   "Form2"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   ScaleHeight     =   7965
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtmutu 
      Height          =   375
      Left            =   10560
      TabIndex        =   28
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cbokode 
      Height          =   315
      Left            =   1800
      TabIndex        =   26
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox cbonpm 
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H8000000B&
      Height          =   1095
      Left            =   240
      TabIndex        =   19
      Top             =   6720
      Width           =   6855
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
         TabIndex        =   24
         Top             =   360
         Width           =   1215
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
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
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtna 
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txttt 
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtuap 
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtuts 
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtquiz 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtuas 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtnamamakul 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5106
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label6 
      Caption         =   "Nilai Mutu"
      Height          =   375
      Left            =   10560
      TabIndex        =   27
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Nilai Akhir"
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "TT"
      Height          =   495
      Left            =   8760
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "UAP"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "UTS"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "QUIS"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "UAS"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nama Makul"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode Makul"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "NPM"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
End
Attribute VB_Name = "frmnilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcari_Click()
Dim rscari As New ADODB.Recordset
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from nilai where kode_matkul='" & txtkodemakul.Text & "'", cn, adOpenStatic, adLockOptimistic
If rscari.RecordCount > 0 Then
    'MsgBox "Data Ada"
    txt = rscari.Fields("nama_mhs").Value
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

Private Sub cbonpm_Click()
Dim rsmahasiswa As New ADODB.Recordset
If rsmahasiswa.State = 1 Then rsmahasiswa.Close
rsmahasiswa.Open "select * from mahasiswa where npm='" & cbonpm & "'", cn, adOpenStatic, adLockOptimistic
txtnama = rsmahasiswa.Fields("nama_mhs").Value
End Sub
Private Sub cbokode_Click()
Dim rsmatkul As New ADODB.Recordset
If rsmatkul.State = 1 Then rsmatkul.Close
rsmatkul.Open "select * from matkul where kode_matkul='" & cbokode & "'", cn, adOpenStatic, adLockOptimistic
txtnamamakul = rsmatkul.Fields("nama_matkul").Value
End Sub
Private Sub isicombo()
Dim rscombo As New ADODB.Recordset
If rscombo.State = 1 Then rscombo.Close
rscombo.Open "select * from matkul", cn, adOpenStatic, adLockOptimistic
Do
    cbokode.AddItem rscombo.Fields("kode_matkul").Value
    rscombo.MoveNext
Loop Until rscombo.EOF
If rscombo.State = 1 Then rscombo.Close
rscombo.Open "select * from mahasiswa", cn, adOpenStatic, adLockOptimistic
Do
    cbonpm.AddItem rscombo.Fields("npm").Value
    rscombo.MoveNext
Loop Until rscombo.EOF
End Sub
Private Sub Form_Load()
If cn.State = 0 Then bukakoneksi
isigrid
isicombo
txtnama.Enabled = False
    cmdedit.Enabled = False
    cmdhapus.Enabled = False
    cmdtambah.Enabled = True
End Sub
Public Sub isigrid()
Dim rsnilai As New ADODB.Recordset
If rsnilai.State = 1 Then rsnilai.Close
rsnilai.Open "select * from nilai", cn, adOpenStatic, adLockOptimistic
If rsnilai.RecordCount > 0 Then
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
    .ColumnHeaders.Add 3, , "UAS"
    .ColumnHeaders.Add 4, , "QUIS"
    .ColumnHeaders.Add 5, , "UTS"
    .ColumnHeaders.Add 6, , "UAP"
    .ColumnHeaders.Add 7, , "TT"
    .ColumnHeaders.Add 8, , "Nilai Mutu"
    .ColumnHeaders(1).Width = 1200
    .ColumnHeaders(2).Width = 2500
    .ColumnHeaders(3).Width = 1200
    .ColumnHeaders(4).Width = 1200
    .ColumnHeaders(5).Width = 1200
    .ColumnHeaders(6).Width = 2000
    .ColumnHeaders(7).Width = 2000
    .ColumnHeaders(8).Width = 2000
  Do Until rsnilai.EOF
    .ListItems.Add 1, , rsnilai.Fields("NPM").Value & ""
    .ListItems(1).SubItems(1) = rsnilai.Fields("nama_mhs").Value & ""
    .ListItems(1).SubItems(2) = rsnilai.Fields("nilai_uas").Value & ""
    .ListItems(1).SubItems(3) = rsnilai.Fields("nilai_quiz").Value & ""
    .ListItems(1).SubItems(4) = rsnilai.Fields("nilai_uts").Value & ""
    .ListItems(1).SubItems(5) = rsnilai.Fields("nilai_uap").Value & ""
    .ListItems(1).SubItems(5) = rsnilai.Fields("nilai_tt").Value & ""
    rsnilai.MoveNext
  Loop
End With
End Sub
Private Sub cmdedit_Click()
If cmdedit.Caption = "&Edit" Then
    cmdedit.Caption = "&Batal"
    cmdsimpan.Enabled = True
    cmdhapus.Enabled = False
    cbokode.Enabled = True
txtnamamakul.Enabled = True
   cbonpm.Enabled = True
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
cn.Execute "DELETE FROM nilai WHERE npm='" & txtnpm & "'"
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
If cbonpm.Text = "" Then
        MsgBox "npm belum diisi"
        cbonpm.SetFocus
ElseIf txtnama.Text = "" Then
        MsgBox "nama belum diisi"
        txtnama.SetFocus
Else
 
    If cmdedit.Caption = "&Batal" Then
        If MsgBox("Data mau diedit ?", vbYesNo) = vbYes Then
            cn.Execute "UPDATE nilai SET npm='" & txtnpm _
            & "',nilai_uas='" & txtuas & "',nilai_quiz='" & txtquiz & "',nilai_uts='" & txtuts & "',nilai_uap'" & txtuap & "',nilai_tt'" & txttt _
            & "' WHERE npm='" & txtnpm & "'"
            isigrid
            cmdedit_Click
        End If
    ElseIf cmdtambah.Caption = "&Batal" Then
        If MsgBox("Data mau ditambah ?", vbYesNo) = vbYes Then
            cn.Execute "INSERT INTO nilai (NPM,nilai_uas,nilai_quiz,nilai_uts,nilai_uap,nilai_tt,nilai_akhir) VALUES ('" & txtnpm _
            & "','" & txtuas & "','" & txtquiz & "','" & txtuts & "','" & txtuap & "','" & txttt & "','" & txtna & "')"
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
  dtanggal.Text = ListView1.SelectedItem.SubItems(4)
  txtalamat.Text = ListView1.SelectedItem.SubItems(5)
End Sub
Private Sub nilai()
Dim hasilakhir As Double
Dim quiz As Integer
Dim uts As Integer
Dim tugas As Integer
Dim uas As Integer
txtna.Text = Val(txtquiz) * 15 / 100 + Val(txttt) * 10 / 100 + Val(txtuts) * 30 / 100 + Val(txtuas) * 30 / 100 + Val(txtuap) * 15 / 100
       If txtna.Text > 75.4 Then
            txtmutu.Text = "A"
        ElseIf txtna.Text = 65.5 - 75.4 Then
            txtmutu.Text = "B"
        ElseIf txtna.Text = 55 - 65.4 Then
            txtmutu.Text = "C"
        ElseIf txtna.Text = 45 - 54.9 Then
            txtmutu.Text = "D"
        Else
            txtmutu.Text = "E"
        End If
hasilakhir = Val(txtna.Text)
tugas = Val(txttt.Text)
uts = Val(txtuts.Text)
uas = Val(txtuas.Text)
hasilakhir = Val(txtna.Text)
hasilmutu = txtmutu.Text
quiz = Val(txtquiz.Text)
End Sub

Private Sub txtna_Change()
nilai
End Sub

Private Sub txtmutu_Change()
nilai
End Sub

Private Sub txtquiz_Change()
nilai
End Sub

Private Sub txttt_Change()
nilai
End Sub

Private Sub txtuas_Change()
nilai
End Sub

Private Sub txtuts_Change()
nilai
End Sub





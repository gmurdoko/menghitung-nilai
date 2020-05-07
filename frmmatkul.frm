VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmatkul 
   Caption         =   "..::::::Data Matakuliah"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.CommandButton cmdcari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H8000000B&
         Height          =   1095
         Left            =   1560
         TabIndex        =   8
         Top             =   2880
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtsks 
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtnama 
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtkode 
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SKS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Matakuliah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode Matakuliah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmmatkul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcari_Click()
Dim rscari As New ADODB.Recordset
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from matkul where kode_matkul='" & txtkode.Text & "'", cn, adOpenStatic, adLockOptimistic
If rscari.RecordCount > 0 Then
    'MsgBox "Data Ada"
    txtnama = rscari.Fields("nama_matkul").Value
    txtsks = rscari.Fields("sks").Value
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
txtsks.Enabled = False
    cmdedit.Enabled = False
    cmdhapus.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = True
End Sub
Public Sub isigrid()
Dim rsmatkul As New ADODB.Recordset
If rsmatkul.State = 1 Then rsmatkul.Close
rsmatkul.Open "select * from matkul", cn, adOpenStatic, adLockOptimistic
If rsmatkul.RecordCount > 0 Then
Else

End If

With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .View = lvwReport
    .LabelEdit = lvwManual
    .ListItems.Clear
    .ColumnHeaders.Add 1, , "Kode Matakuliah"
    .ColumnHeaders.Add 2, , "Nama Matakuliah"
    .ColumnHeaders.Add 3, , "SKS"
    .ColumnHeaders(1).Width = 1500
    .ColumnHeaders(2).Width = 2500
    .ColumnHeaders(3).Width = 1200
  Do Until rsmatkul.EOF
    .ListItems.Add 1, , rsmatkul.Fields("kode_matkul").Value & ""
    .ListItems(1).SubItems(1) = rsmatkul.Fields("nama_matkul").Value & ""
    .ListItems(1).SubItems(2) = rsmatkul.Fields("sks").Value & ""
    rsmatkul.MoveNext
  Loop
End With
End Sub
Private Sub cmdedit_Click()
If cmdedit.Caption = "&Edit" Then
    cmdedit.Caption = "&Batal"
    cmdsimpan.Enabled = True
    cmdhapus.Enabled = False
    txtkode.Enabled = True
    txtnama.Enabled = True
    txtsks.Enabled = True
Else
   cmdedit.Caption = "&Edit"
   Form_Load
End If
End Sub
Private Sub cmdhapus_Click()
If MsgBox("Data mau dihapus ?", vbYesNo) = vbYes Then
cn.Execute "DELETE FROM matkul WHERE kode_matkul='" & txtkode & "'"
isigrid
End If
txtkode.Text = ""
txtnama.Text = ""
txtsks.Text = ""
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
If txtkode.Text = "" Then
        MsgBox "kode belum diisi"
        txtnpm.SetFocus
ElseIf txtnama.Text = "" Then
        MsgBox "nama belum diisi"
        txtnama.SetFocus
Else
 
    If cmdedit.Caption = "&Batal" Then
        If MsgBox("Data mau diedit ?", vbYesNo) = vbYes Then
            cn.Execute "UPDATE matkul SET nama_matkul='" & txtnama & "',sks='" & txtsks _
            & "' WHERE kode_matkul='" & txtkode & "'"
            isigrid
            cmdedit_Click
        End If
    ElseIf cmdtambah.Caption = "&Batal" Then
        If MsgBox("Data mau ditambah ?", vbYesNo) = vbYes Then
            cn.Execute "INSERT INTO matkul (kode_matkul,nama_matkul,sks) VALUES ('" & txtkode _
            & "','" & txtnama & "','" & txtsks & "')"
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
    txtkode.Enabled = True
    txtnama.Enabled = True
    txtsks.Enabled = True
Else
cmdtambah.Caption = "&Tambah"
Form_Load
End If
End Sub

Private Sub ListView1_DblClick()
 txtkode.Text = ListView1.SelectedItem.Text
  txtnama.Text = ListView1.SelectedItem.SubItems(1)
  txtsks.Text = ListView1.SelectedItem.SubItems(2)
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = False
End Sub

Private Sub txtnpm_LostFocus()
cmdcari_Click
End Sub

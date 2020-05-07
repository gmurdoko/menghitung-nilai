VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdosen 
   Caption         =   "..::::::Data Dosen"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H8000000B&
         Height          =   1095
         Left            =   1560
         TabIndex        =   13
         Top             =   4200
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdcari 
         Caption         =   "&Cari"
         Height          =   375
         Left            =   7320
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   5400
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComCtl2.DTPicker tanggald 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   2640
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   873
         _Version        =   393216
         Format          =   106364929
         CurrentDate     =   42534
      End
      Begin VB.TextBox txtalamatd 
         Height          =   615
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txttempatd 
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtnamad 
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtnip 
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
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
         TabIndex        =   5
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Lahir"
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
         TabIndex        =   4
         Top             =   2760
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
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
         Top             =   2040
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Dosen"
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
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
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
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmdosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcari_Click()
Dim rscari As New ADODB.Recordset
If rscari.State = 1 Then rscari.Close
rscari.Open "select * from dosen where nip='" & txtnip.Text & "'", cn, adOpenStatic, adLockOptimistic
If rscari.RecordCount > 0 Then
    'MsgBox "Data Ada"
    txtnamad = rscari.Fields("nama_dosen").Value
    txttempatd = rscari.Fields("tempat_lahir").Value
    tanggald = rscari.Fields("tanggal_lahir").Value
    txtalamatd = rscari.Fields("alamat").Value
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
txtnamad.Enabled = False
txttempatd.Enabled = False
tanggald.Enabled = False
txtalamatd.Enabled = False
    cmdedit.Enabled = False
    cmdhapus.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = True
End Sub
Public Sub isigrid()
Dim rsdosen As New ADODB.Recordset
If rsdosen.State = 1 Then rsdosen.Close
rsdosen.Open "select * from dosen", cn, adOpenStatic, adLockOptimistic
If rsdosen.RecordCount > 0 Then
Else

End If

With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .View = lvwReport
    .LabelEdit = lvwManual
    .ListItems.Clear
    .ColumnHeaders.Add 1, , "NIP"
    .ColumnHeaders.Add 2, , "Nama Dosen"
    .ColumnHeaders.Add 3, , "Tempat Lahir"
    .ColumnHeaders.Add 4, , "Tanggal Lahir"
    .ColumnHeaders.Add 5, , "Alamat"
    .ColumnHeaders(1).Width = 1500
    .ColumnHeaders(2).Width = 2500
    .ColumnHeaders(3).Width = 1500
    .ColumnHeaders(4).Width = 1200
    .ColumnHeaders(5).Width = 2000
  Do Until rsdosen.EOF
    .ListItems.Add 1, , rsdosen.Fields("nip").Value & ""
    .ListItems(1).SubItems(1) = rsdosen.Fields("nama_dosen").Value & ""
    .ListItems(1).SubItems(2) = rsdosen.Fields("tempat_lahir").Value & ""
    .ListItems(1).SubItems(3) = rsdosen.Fields("tanggal_lahir").Value & ""
    .ListItems(1).SubItems(4) = rsdosen.Fields("alamat").Value & ""
    rsdosen.MoveNext
  Loop
End With
End Sub
Private Sub cmdedit_Click()
If cmdedit.Caption = "&Edit" Then
    cmdedit.Caption = "&Batal"
    cmdsimpan.Enabled = True
    cmdhapus.Enabled = False
    txtnpm.Enabled = True
txtnamad.Enabled = True
    txttempatd.Enabled = True
    tanggald.Enabled = True
    txtalamatd.Enabled = True
Else
   cmdedit.Caption = "&Edit"
   Form_Load
End If
End Sub
Private Sub cmdhapus_Click()
If MsgBox("Data mau dihapus ?", vbYesNo) = vbYes Then
cn.Execute "DELETE FROM dosen WHERE nip='" & txtnip & "'"
isigrid
End If
txtnip.Text = ""
txtnamad.Text = ""
tanggald = ""
txttempatd.Text = ""
txtalamatd.Text = ""
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
If txtnip.Text = "" Then
        MsgBox "nip belum diisi"
        txtnip.SetFocus
ElseIf txtnamad.Text = "" Then
        MsgBox "nama belum diisi"
        txtnamad.SetFocus
Else
 
    If cmdedit.Caption = "&Batal" Then
        If MsgBox("Data mau diedit ?", vbYesNo) = vbYes Then
            cn.Execute "UPDATE dosen SET nama_dosen='" & txtnamad _
            & "',tempat_lahir='" & txttempatd & "',tanggal_lahir='" & tanggald & "',alamat='" & txtalamat _
            & "' WHERE nip='" & txtnip & "'"
            isigrid
            cmdedit_Click
        End If
    ElseIf cmdtambah.Caption = "&Batal" Then
        If MsgBox("Data mau ditambah ?", vbYesNo) = vbYes Then
            cn.Execute "INSERT INTO dosen (nip,nama_dosen,tempat_lahir,tanggal_lahir,alamat) VALUES ('" & txtnip _
            & "','" & txtnamad & "','" & txttempatd & "','" & tanggald & "','" & txtalamatd & "')"
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
txtnip.Enabled = True
txtnamad.Enabled = True
    txttempatd.Enabled = True
    tanggald.Enabled = True
    txtalamatd.Enabled = True
Else
cmdtambah.Caption = "&Tambah"
Form_Load
End If
End Sub

Private Sub ListView1_DblClick()
 txtnip.Text = ListView1.SelectedItem.Text
  txtnamad.Text = ListView1.SelectedItem.SubItems(1)
  txttempatd.Text = ListView1.SelectedItem.SubItems(2)
  tanggald = ListView1.SelectedItem.SubItems(3)
  txtalamatd.Text = ListView1.SelectedItem.SubItems(4)
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = False
End Sub

Private Sub txtnip_LostFocus()
cmdcari_Click
End Sub

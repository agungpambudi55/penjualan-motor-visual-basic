VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP MOTOR"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPFrame XPFrame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      Caption         =   "Entri Pengguna"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPControls.XPButton XPButton8 
         Height          =   735
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "Menghapus isi pada teks"
         Top             =   2520
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Picture         =   "Form10.frx":0442
         Caption         =   "Refresh"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton XPButton7 
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "Pencarian yang menggunakan e-mail"
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "Pencarian"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton XPButton6 
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         ToolTipText     =   "Merefresh / meperbarui data"
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Caption         =   "Refresh Data"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form10.frx":2244
         Height          =   1575
         Left            =   2760
         TabIndex        =   17
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Email"
            Caption         =   "Email"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nama"
            Caption         =   "Nama"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Kode"
            Caption         =   "Kode"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Status"
            Caption         =   "Status"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
            EndProperty
         EndProperty
      End
      Begin XPControls.XPButton XPButton5 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Membatalkan data yang akan diedit"
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Picture         =   "Form10.frx":2259
         Caption         =   "Batal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton XPButton4 
         Height          =   735
         Left            =   4560
         TabIndex        =   13
         ToolTipText     =   "kembali ke halaman awal"
         Top             =   3360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Picture         =   "Form10.frx":405B
         Caption         =   "Home"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton XPButton3 
         Height          =   735
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "Menghapus data"
         Top             =   3360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Enabled         =   0   'False
         Picture         =   "Form10.frx":5E5D
         Caption         =   "Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   405
         Left            =   840
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "User"
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   840
         MaxLength       =   40
         TabIndex        =   1
         ToolTipText     =   "Masukkan e-mail yang digunakan"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Masukkan kode untuk membuka"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         MaxLength       =   40
         TabIndex        =   2
         ToolTipText     =   "Masukkan nama, terserah mau panjang atau pendek"
         Top             =   960
         Width           =   1815
      End
      Begin XPControls.XPButton XPButton2 
         Height          =   735
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "Memperbarui / mengupdate data"
         Top             =   3360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Enabled         =   0   'False
         Picture         =   "Form10.frx":7B97
         Caption         =   "Dibarui"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin XPControls.XPButton XPButton1 
         Height          =   735
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Menyimpan data"
         Top             =   2520
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Picture         =   "Form10.frx":9999
         Caption         =   "Tambah"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2640
         Top             =   3720
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Penjualan Motor"
         OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Penjualan Motor"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from pengguna"
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
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   2760
         Picture         =   "Form10.frx":B79B
         Top             =   2400
         Width           =   2460
      End
      Begin VB.Label Label4 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Adodc1.RecordSource = "Select * from pengguna where Email='" & Text1 & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc1.Recordset
Text1 = !Email
Text2 = !Nama
Text3 = !Kode
Text4 = !Status
End With
XPButton2.Enabled = True
XPButton3.Enabled = True
XPButton1.Enabled = False
XPButton5.Visible = True
XPButton8.Enabled = False
Text1.SetFocus
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
End If
End If
End Sub

Private Sub XPButton1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
D = "select * from pengguna where Email='" & Text1 & "'"
Adodc1.RecordSource = D
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Email Yang Anda Masukkan Sudah Ada", vbInformation, "Informasi"
Else
Adodc1.RecordSource = "select *from pengguna where Email='" & Text1 & "'"
Adodc1.Refresh
With Adodc1.Recordset
.AddNew
!Email = Text1
!Nama = Text2
!Kode = Text3
!Status = Text4
.Update
End With
MsgBox "Data Tersimpan", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
Text1.SetFocus
Adodc1.RecordSource = "select* from pengguna"
Adodc1.Refresh
End If
End If
End Sub

Private Sub XPButton2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
Adodc1.RecordSource = "select *from pengguna where Email='" & Text1 & "'"
If Adodc1.Recordset.RecordCount > 0 Then
With Adodc1.Recordset
!Nama = Text2
!Kode = Text3
!Status = Text4
.Update
End With
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton5.Visible = False
XPButton8.Enabled = True
MsgBox "Data Tersimpan", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
Text1.SetFocus
Adodc1.RecordSource = "select* from pengguna"
Adodc1.Refresh
End If
End If
End Sub

Private Sub XPButton3_Click()
A = MsgBox("Apakah Akan Dihapus?", vbQuestion + vbYesNo, "Hapus Data")
If A = vbYes Then
Adodc1.Recordset.Delete
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton5.Visible = False
XPButton8.Enabled = True
MsgBox "Data Telah Terhapus", vbInformation, "Information"
Text1.SetFocus
Adodc1.RecordSource = "select* from pengguna"
Adodc1.Refresh
End If
End Sub

Private Sub XPButton4_Click()
Form6.XPButton8.Enabled = True
Form6.Show
Unload Me
End Sub

Private Sub XPButton5_Click()
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton5.Visible = False
XPButton8.Enabled = True
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
Text1.SetFocus
End Sub

Private Sub XPButton6_Click()
Adodc1.RecordSource = "select* from pengguna"
Adodc1.Refresh
End Sub

Private Sub XPButton7_Click()
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton5.Visible = False
XPButton8.Enabled = True
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
Text1.SetFocus
C = InputBox("Masukkan E-mailnya!", "Pencarian")
D = "select * from pengguna where Email='" & C & "'"
Adodc1.RecordSource = D
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
Text1 = Adodc1.Recordset!Email
Text2 = Adodc1.Recordset!Nama
Text3 = Adodc1.Recordset!Kode
Text4 = Adodc1.Recordset!Status
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
Adodc1.RecordSource = "select* from pengguna"
Adodc1.Refresh
End If
If E = vbYes Then
XPButton2.Enabled = True
XPButton3.Enabled = True
XPButton1.Enabled = False
XPButton5.Visible = True
XPButton8.Enabled = False
End If
End Sub

Private Sub XPButton8_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = "User"
End Sub


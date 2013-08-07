VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP MOTOR"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPFrame XPFrame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      Caption         =   "Log In"
      ForeColor       =   0
      CaptionAlignment=   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPControls.XPButton XPButton3 
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         ToolTipText     =   "Informasi Log in"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         Picture         =   "Form9.frx":0442
         Caption         =   "Info"
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
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Masukkan kode / password anda"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Masukkan nama anda"
         Top             =   600
         Width           =   1695
      End
      Begin XPControls.XPButton XPButton2 
         Height          =   735
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Keluar dari software AP MOTOR"
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Picture         =   "Form9.frx":0E8C
         Caption         =   "Keluar"
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
         TabIndex        =   1
         ToolTipText     =   "Masuk ke software AP MOTOR "
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         Enabled         =   0   'False
         Picture         =   "Form9.frx":29DE
         Caption         =   "Log in"
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
         Left            =   2040
         Top             =   840
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
         Height          =   525
         Left            =   2760
         Picture         =   "Form9.frx":47E0
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
If Text1 = "" Or Text2 = "" Then
XPButton1.Enabled = False
Else
XPButton1.Enabled = True
End If
End Sub

Private Sub Text2_Change()
If Text1 = "" Or Text2 = "" Then
XPButton1.Enabled = False
Else
XPButton1.Enabled = True
End If
End Sub



Private Sub XPButton1_Click()
A = "select * from pengguna where Nama='" & Text1 & "' and Kode='" & Text2 & "'"
Adodc1.RecordSource = A
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Selamat Datang Di AP MOTOR", vbInformation, "Welcome"
    If Adodc1.Recordset!Status = "Admin" Then
    Form6.XPButton8.Enabled = True
    Form6.XPFrame3.BackColor = &H80FF80
    Form6.Label4.Caption = "Anda bisa membuat atau mengentri data pengguna"
    Else
    Form6.Label4.Caption = "Yang Bisa Entri Pengguna             Hanya Admin"
    End If
Unload Me
Form8.Show
Else
MsgBox "Anda Belum Terdaftar!", vbInformation, "Informasi"
End If

End Sub

Private Sub XPButton2_Click()
End
End Sub

Private Sub XPButton3_Click()
MsgBox "Yang bisa menggunakan software ini orang yang sudah terdaftar dalam user penggunaan. Sedangkan yang bisa membuat atau mengentri data pengguna software ini hanya admin", vbInformation, "Informasi"

End Sub

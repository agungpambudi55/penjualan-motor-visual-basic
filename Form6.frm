VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP MOTOR"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8895
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPFrame XPFrame3 
      Height          =   2415
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4260
      BackColor       =   14737632
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPControls.XPButton XPButton8 
         Height          =   1095
         Left            =   960
         TabIndex        =   15
         ToolTipText     =   "Mengentri data pengguna pada login"
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1931
         Enabled         =   0   'False
         Picture         =   "Form6.frx":0442
         Caption         =   "Enti pengguna"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureAlignment=   4
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
   End
   Begin XPControls.XPButton XPButton11 
      Height          =   975
      Left            =   3960
      TabIndex        =   13
      ToolTipText     =   "keluar ke log in"
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      Picture         =   "Form6.frx":48EC
      Caption         =   "Log out"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
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
      Left            =   7800
      Top             =   8040
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
   Begin XPControls.XPButton XPButton9 
      Height          =   1815
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Pencipta software ini"
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   3201
      Picture         =   "Form6.frx":793E
      Caption         =   ""
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
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin XPControls.XPFrame XPFrame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   8493
      BackColor       =   8421504
      BackStyle       =   0
      Caption         =   "ENTRI DATA"
      ForeColor       =   0
      CaptionAlignment=   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Mengentri data motor"
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":F658
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton2 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Mengentri data karyawan"
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":14BD6
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton3 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Mengentri data transaksi pembelian"
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":1A154
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton4 
      Height          =   1215
      Left            =   7320
      TabIndex        =   4
      ToolTipText     =   "Menampilkan laporan motor"
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":1F6D2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton5 
      Height          =   1215
      Left            =   7320
      TabIndex        =   5
      ToolTipText     =   "Menampilkan laporan karyawan"
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":24C50
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPButton XPButton6 
      Height          =   1215
      Left            =   7320
      TabIndex        =   6
      ToolTipText     =   "Menampilkan laporan transaksi pembelian"
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      Picture         =   "Form6.frx":2A1CE
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   4
   End
   Begin XPControls.XPFrame XPFrame2 
      Height          =   4815
      Left            =   7200
      TabIndex        =   7
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   8493
      BackColor       =   8421504
      BackStyle       =   0
      Caption         =   "DATA REPORT"
      ForeColor       =   0
      CaptionAlignment=   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton10 
      Height          =   1815
      Left            =   7320
      TabIndex        =   12
      ToolTipText     =   "Pencipta software ini"
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   3201
      Picture         =   "Form6.frx":2F74C
      Caption         =   ""
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form6.frx":37466
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   17
      Top             =   1440
      Width           =   5055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      X1              =   0
      X2              =   13200
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Index           =   1
      Left            =   0
      Picture         =   "Form6.frx":37523
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   9000
   End
   Begin VB.Image Image2 
      Height          =   7215
      Index           =   0
      Left            =   0
      Picture         =   "Form6.frx":3E30C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Polorejo-Babadan-Ponorogo-Jawa Timur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jalan Letjen Suprapto Sukowati No.49 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AP MOTOR===>  MURAH, NYAMAN, DAN BERKELAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6480
      TabIndex        =   8
      Top             =   0
      Width           =   13815
   End
   Begin VB.Image Image2 
      Height          =   7095
      Index           =   2
      Left            =   7080
      Picture         =   "Form6.frx":45E4A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Form6.frx":4D988
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13200
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 200
If Label1.Left < -12000 Then
Label1.Left = 11000
End If
If Label5.Visible = True Then
Label5.Visible = False
ElseIf Label5.Visible = False Then
Label5.Visible = True
End If
End Sub

Private Sub XPButton1_Click()
Form1.Show
Unload Me
End Sub

Private Sub XPButton10_Click()
MsgBox "Software ini diciptakan oleh Pemilik AP MOTOR yaitu Agung Pambudi", vbInformation, "Creator"
End Sub

Private Sub XPButton11_Click()
XPButton8.Enabled = False
XPFrame3.BackColor = &HE0E0E0
Label4.Caption = "Yang Bisa Entri Pengguna             Hanya Admin"
Form9.Show
Unload Me
End Sub

Private Sub XPButton2_Click()
Form2.Show
Unload Me
End Sub

Private Sub XPButton3_Click()
Form3.Show
Unload Me
End Sub

Private Sub XPButton4_Click()
DataReport1.Show
DataReport1.Refresh
End Sub

Private Sub XPButton5_Click()
DataReport2.Show
DataReport2.Refresh
End Sub

Private Sub XPButton6_Click()
DataReport3.Show
DataReport3.Refresh
End Sub




Private Sub XPButton8_Click()
Form10.Show
Unload Me
End Sub

Private Sub XPButton9_Click()
MsgBox "Software ini diciptakan oleh Pemilik AP MOTOR yaitu Agung Pambudi", vbInformation, "Creator"
End Sub


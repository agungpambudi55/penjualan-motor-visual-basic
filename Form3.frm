VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP MOTOR"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   8910
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ENTRI DATA"
      TabPicture(0)   =   "Form3.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label18"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label19"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "XPButton6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "XPButton1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "XPButton2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "XPButton3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "XPButton4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Adodc1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "XPButton5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "XPFrame1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "XPFrame4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text4"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "DTPicker1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text5"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text6"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text7"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text8"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text9"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text10"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "XPButton10"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "XPButton11"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Adodc3"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "LIHAT DATA"
      TabPicture(1)   =   "Form3.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "XPFrame6"
      Tab(1).Control(1)=   "XPFrame5"
      Tab(1).Control(2)=   "XPButton9"
      Tab(1).Control(3)=   "XPButton8"
      Tab(1).Control(4)=   "XPButton7"
      Tab(1).Control(5)=   "DataGrid1"
      Tab(1).Control(6)=   "Adodc2"
      Tab(1).Control(7)=   "XPButton12"
      Tab(1).Control(8)=   "XPButton13"
      Tab(1).Control(9)=   "XPButton14"
      Tab(1).Control(10)=   "XPButton15"
      Tab(1).Control(11)=   "XPButton16"
      Tab(1).ControlCount=   12
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   7200
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         RecordSource    =   "select * from motor"
         Caption         =   "Adodc3"
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
      Begin XPControls.XPButton XPButton11 
         Height          =   375
         Left            =   8280
         TabIndex        =   50
         ToolTipText     =   "Lihat data motor"
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         Caption         =   "^"
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
      Begin XPControls.XPButton XPButton10 
         Height          =   375
         Left            =   5640
         TabIndex        =   49
         ToolTipText     =   "Lihat data karyawan"
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         Caption         =   "^"
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
      Begin VB.TextBox Text10 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   39
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   32
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         TabIndex        =   31
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   6960
         MaxLength       =   15
         TabIndex        =   8
         ToolTipText     =   "Masukkan nomor mesin"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Masukkan nomor karyawan"
         Top             =   1200
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Ubah tanggal pembelian motor"
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   7602177
         CurrentDate     =   41015
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Masukkan nomor antrian / nomor nota"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Masukkan nama pembeli"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Masukkan alamat pembeli"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Masukkan nomor telephone hp/rumah"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form3.frx":047A
         Left            =   1320
         List            =   "Form3.frx":0484
         TabIndex        =   4
         ToolTipText     =   "Masukkan jenis kelamin pembeli"
         Top             =   2160
         Width           =   1575
      End
      Begin XPControls.XPFrame XPFrame4 
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Top             =   5640
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BackColor       =   12632256
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
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "     Jl. Letj. Suprapto Sukowati No. 49 Babadan Ponorogo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   20
            Top             =   0
            Width           =   6015
         End
      End
      Begin XPControls.XPFrame XPFrame1 
         Height          =   855
         Left            =   6600
         TabIndex        =   21
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1508
         BackColor       =   16761024
         BackStyle       =   0
         Caption         =   "Time and Date"
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
      Begin XPControls.XPButton XPButton5 
         Height          =   975
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "Mengedit data"
         Top             =   4320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":049E
         Caption         =   "Edit"
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
         Left            =   6000
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         RecordSource    =   "select * from motor"
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
      Begin XPControls.XPButton XPButton4 
         Height          =   975
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "Menghapus isian pada teks"
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":34F0
         Caption         =   "Refresh"
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
      Begin XPControls.XPButton XPButton3 
         Height          =   975
         Left            =   1920
         TabIndex        =   11
         ToolTipText     =   "Memperbarui/mengupdate data"
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form3.frx":6542
         Caption         =   "Dibarui"
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
      Begin XPControls.XPButton XPButton2 
         Height          =   975
         Left            =   4080
         TabIndex        =   12
         ToolTipText     =   "Menghapus data"
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form3.frx":9594
         Caption         =   "Hapus"
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
      Begin XPControls.XPButton XPButton1 
         Height          =   975
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Menyimpan data"
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":C5E6
         Caption         =   "Simpan"
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
      Begin XPControls.XPButton XPButton6 
         Height          =   975
         Left            =   3000
         TabIndex        =   14
         ToolTipText     =   "Membatalkan mengedit"
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":F638
         Caption         =   "Batal"
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
      Begin XPControls.XPFrame XPFrame6 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   41
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2566
         BackStyle       =   0
         Caption         =   "Pencarian Menggunakan :"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPFrame XPFrame5 
         Height          =   495
         Left            =   -74760
         TabIndex        =   42
         Top             =   5160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         BackColor       =   16761024
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
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   44
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data ="
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
            TabIndex        =   43
            Top             =   120
            Width           =   1215
         End
      End
      Begin XPControls.XPButton XPButton9 
         Height          =   975
         Left            =   -72000
         TabIndex        =   45
         ToolTipText     =   "Menghapus data"
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form3.frx":1268A
         Caption         =   "Hapus"
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
      Begin XPControls.XPButton XPButton8 
         Height          =   975
         Left            =   -74640
         TabIndex        =   46
         ToolTipText     =   "Pencarian dengan menggunakan nomor faktur"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":156DC
         Caption         =   "Nomor Faktur"
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
      Begin XPControls.XPButton XPButton7 
         Height          =   975
         Left            =   -70560
         TabIndex        =   47
         ToolTipText     =   "Merefresh / memperbarui data"
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":1872E
         Caption         =   "Refresh Data"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form3.frx":1B780
         Height          =   2775
         Left            =   -74760
         TabIndex        =   48
         Top             =   1920
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   16777215
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "No_Faktur"
            Caption         =   "No_Faktur"
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
            DataField       =   "Alamat"
            Caption         =   "Alamat"
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
            DataField       =   "Jenis_Kelamin"
            Caption         =   "Jenis_Kelamin"
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
         BeginProperty Column04 
            DataField       =   "No_Telephone"
            Caption         =   "No_Telephone"
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
         BeginProperty Column05 
            DataField       =   "Tanggal_Beli"
            Caption         =   "Tanggal_Beli"
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
         BeginProperty Column06 
            DataField       =   "No_Karyawan"
            Caption         =   "No_Karyawan"
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
         BeginProperty Column07 
            DataField       =   "Nama_Karyawan"
            Caption         =   "Nama_Karyawan"
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
         BeginProperty Column08 
            DataField       =   "No_Mesin"
            Caption         =   "No_Mesin"
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
         BeginProperty Column09 
            DataField       =   "Merk"
            Caption         =   "Merk"
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
         BeginProperty Column10 
            DataField       =   "Type"
            Caption         =   "Type"
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
         BeginProperty Column11 
            DataField       =   "Harga"
            Caption         =   "Harga"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   -74760
         Top             =   4680
         Width           =   8160
         _ExtentX        =   14393
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
         RecordSource    =   "select * from transaksi_pembeli"
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
         _Version        =   393216
      End
      Begin XPControls.XPButton XPButton12 
         Height          =   975
         Left            =   -72600
         TabIndex        =   51
         ToolTipText     =   "Pencarian dengan menggunakan nama pembeli"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":1B795
         Caption         =   "Nama Pembeli"
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
      Begin XPControls.XPButton XPButton13 
         Height          =   975
         Left            =   -71400
         TabIndex        =   52
         ToolTipText     =   "Pencarian dengan menggunakan jenis kelamin pembeli"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":1E7E7
         Caption         =   "Jenis Kelamin"
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
      Begin XPControls.XPButton XPButton14 
         Height          =   975
         Left            =   -70200
         TabIndex        =   53
         ToolTipText     =   "Pencarian dengan menggunakan tanggal beli"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":21839
         Caption         =   "Tanggal Beli"
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
      Begin XPControls.XPButton XPButton15 
         Height          =   975
         Left            =   -69000
         TabIndex        =   54
         ToolTipText     =   "Pencarian dengan menggunakan nomor karyawan"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":2488B
         Caption         =   "Nomor Karyawan"
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
      Begin XPControls.XPButton XPButton16 
         Height          =   975
         Left            =   -67800
         TabIndex        =   55
         ToolTipText     =   "Pencarian dengan menggunakan nomor mesin"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form3.frx":278DD
         Caption         =   "Nomor Mesin"
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
      Begin VB.Label Label19 
         Caption         =   "No. Telephone"
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
         TabIndex        =   40
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Harga"
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
         Left            =   6000
         TabIndex        =   38
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Type"
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
         Left            =   6000
         TabIndex        =   37
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Merk"
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
         Left            =   6000
         TabIndex        =   36
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "No. Mesin"
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
         Left            =   6000
         TabIndex        =   35
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Nama Karyawan"
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
         Left            =   3000
         TabIndex        =   34
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "No. Karyawan"
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
         Left            =   3000
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Faktur"
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
         TabIndex        =   29
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Pembeli"
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
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
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Beli"
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
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   24
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   23
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   1320
         Left            =   5280
         Picture         =   "Form3.frx":2A92F
         Top             =   4320
         Width           =   3105
      End
      Begin VB.Label Label11 
         Caption         =   "Pencarian Dg Menggunakan Keypress"
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
         Left            =   2520
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   1920
      Top             =   7440
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1440
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   960
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   7440
   End
   Begin XPControls.XPFrame XPFrame2 
      Height          =   255
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      BackColor       =   0
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
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "          Nama= Agung Pambudi   No.Absen= 03   Kelas= X RPL A   Sekolah= SMKN1 JENANGAN PONOROGO          "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by :"
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
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI PEMBELIAN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      X1              =   0
      X2              =   9000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   0
      Picture         =   "Form3.frx":37FF1
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   9000
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "Form3.frx":3EDDA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Home 
         Caption         =   "Home"
      End
      Begin VB.Menu Keluar 
         Caption         =   "Keluar"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
Label9.Caption = "          Nama= Agung Pambudi   No.Absen= 03   Kelas= X RPL A   Sekolah= SMKN1 JENANGAN PONOROGO          "
Label10.Caption = "     Jl. Letj. Suprapto Sukowati No. 49 Babadan Ponorogo"
End Sub

Private Sub Home_Click()
Form6.Show
Form6.Visible = True
Unload Me
End Sub

Private Sub Keluar_Click()
i = MsgBox("Apakah Anda ingin Keluar", vbQuestion + vbYesNo, "Pertanyaan")
If i = vbYes Then
End
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Adodc1.RecordSource = "Select * from transaksi_pembeli where No_Faktur='" & Text1 & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc1.Recordset
Text1 = !No_Faktur
Text2 = !Nama
Text3 = !Alamat
Combo1 = !Jenis_Kelamin
Text4 = !No_Telephone
DTPicker1 = !Tanggal_Beli
Text5 = !No_Karyawan
Text6 = !Nama_Karyawan
Text7 = !No_Mesin
Text8 = !Merk
Text9 = !Type
Text10 = !Harga
End With
XPButton2.Enabled = True
XPButton3.Enabled = True
XPButton1.Enabled = False
XPButton4.Enabled = False
XPButton5.Visible = False
XPButton6.Visible = True
Text2.SetFocus
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
End If
End If
If ((KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57) Then
KeyAscii = 0
End If
End Sub

Private Sub Text2_Change()
Dim i As Integer
i = Text2.SelStart
Text2.Text = StrConv(Text2.Text, vbProperCase)
Text2.SelStart = i
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or symbol = "-" Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
KeyAscii = 0
End If
End Sub

Private Sub Text3_Change()
Dim i As Integer
i = Text3.SelStart
Text3.Text = StrConv(Text3.Text, vbProperCase)
Text3.SelStart = i
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If ((KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57) Then
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Adodc2.RecordSource = "Select * from Karyawan where No_Karyawan='" & Text5 & "'"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc2.Recordset
Text5 = !No_Karyawan
Text6 = !Nama_Karyawan
End With
Text7.SetFocus
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
Text5 = ""
Text6 = ""
Text5.SetFocus
End If
End If
If ((KeyAscii < 48 And KeyAscii <> 8) Or KeyAscii > 57) Then
KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Adodc3.RecordSource = "Select * from Motor where No_Mesin='" & Text7 & "'"
Adodc3.Refresh
If Adodc3.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc3.Recordset
Text7 = !No_Mesin
Text8 = !Merk
Text9 = !Type
Text10 = !Harga
End With
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text7.SetFocus
End If
End If
End Sub

Private Sub Timer1_Timer()
Label7.Caption = Time
Label8.Caption = Date
End Sub

Private Sub Timer3_Timer()
Label10.Left = Label10.Left - 100
If Label10.Left < -4000 Then
Label10.Left = 3500
End If
End Sub

Private Sub Timer2_Timer()
Label9.Left = Label9.Left - 100
If Label9.Left < -6800 Then
Label9.Left = 3500
End If
End Sub

Private Sub Timer4_Timer()
Label11.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End Sub

Private Sub Timer5_Timer()
Label21.Caption = Adodc2.Recordset.RecordCount
End Sub

Private Sub XPButton1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
D = "select * from transaksi_pembeli where No_Faktur='" & Text1 & "'"
Adodc1.RecordSource = D
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Tidak Bisa Disimpan, Karena Nomor Faktur Tersebut Sudah Ada", vbInformation, "Informasi"
Else
Adodc1.RecordSource = "select *from transaksi_pembeli where No_Faktur='" & Text1 & "'"
Adodc1.Refresh
With Adodc1.Recordset
.AddNew
!No_Faktur = Text1
!Nama = Text2
!Alamat = Text3
!Jenis_Kelamin = Combo1
!No_Telephone = Text4
!Tanggal_Beli = DTPicker1
!No_Karyawan = Text5
!Nama_Karyawan = Text6
!No_Mesin = Text7
!Merk = Text8
!Type = Text9
!Harga = Text10
.Update
End With
Adodc3.Refresh
Adodc3.Recordset.Delete
MsgBox "Data Tersimpan", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
End If
End If
End Sub

Private Sub XPButton10_Click()
Form5.Show
Form3.Enabled = False
End Sub

Private Sub XPButton11_Click()
Form3.Enabled = False
Form4.Show
End Sub

Private Sub XPButton12_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Nama Pembeli Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where Nama='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If

End Sub

Private Sub XPButton14_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Tanggal Beli dengan format bb/hh/tttt Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where Tanggal_Beli='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If
End Sub

Private Sub XPButton15_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Nomor Karyawan Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where No_Karyawan='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If
End Sub

Private Sub XPButton16_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Nomor Mesin Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where No_Mesin='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If

End Sub

Private Sub XPButton13_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Jenis Kelamin Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where Jenis_Kelamin='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If
End Sub

Private Sub XPButton17_Click()

End Sub

Private Sub XPButton2_Click()
A = MsgBox("Apakah Akan Dihapus?", vbQuestion + vbYesNo, "Hapus Data")
If A = vbYes Then
Adodc1.Recordset.Delete
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
MsgBox "Data Telah Terhapus", vbInformation, "Information"
Text1.SetFocus
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
End If
End Sub

Private Sub XPButton3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
Adodc1.RecordSource = "select *from transaksi_pembeli where No_Faktur='" & Text1 & "'"
If Adodc1.Recordset.RecordCount > 0 Then
With Adodc1.Recordset
!Nama = Text2
!Alamat = Text3
!Jenis_Kelamin = Combo1
!No_Telephone = Text4
!Tanggal_Beli = DTPicker1
!No_Karyawan = Text5
!Nama_Karyawan = Text6
!No_Mesin = Text7
!Merk = Text8
!Type = Text9
!Harga = Text10
.Update
End With
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
MsgBox "Data Tersimpan", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
End If
End If
End Sub

Private Sub XPButton4_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
End Sub

Private Sub XPButton5_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Then
MsgBox "Jika Mau Edit, Masukkan Datanya Dengan Keypress Pada Text Nomor Faktur!", vbInformation, "Informasi"
If KeyAscii = 13 Then
Adodc1.RecordSource = "Select * from transaksi_pembeli where No_Faktur='" & Text1 & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc1.Recordset
!No_Faktur = Text1
!Nama = Text2
!Alamat = Text3
!Jenis_Kelamin = Combo1
!No_Telephone = Text4
!Tanggal_Beli = DTPicker1
!No_Karyawan = Text5
!Nama_Karyawan = Text6
!No_Mesin = Text7
!Merk = Text8
!Type = Text9
!Harga = Text10
End With
XPButton2.Enabled = True
XPButton3.Enabled = True
XPButton1.Enabled = False
XPButton4.Enabled = False
XPButton5.Visible = False
XPButton6.Visible = True
Text2.SetFocus
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
End If
End If
End If
End Sub

Private Sub XPButton6_Click()
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Text8 = ""
Text9 = ""
Text10 = ""
Text1.SetFocus
End Sub

Private Sub XPButton7_Click()
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
XPButton9.Enabled = False
End Sub

Private Sub XPButton8_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Nomor Faktur Yang Akan Anda Cari!", "Pencarian")
D = "select * from transaksi_pembeli where No_Faktur='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
f = MsgBox("Data Yang Dicari Kosong!", vbInformation, "Informasi")
End If
If E = vbYes Then
XPButton9.Enabled = True
Else
If D = vbNo Then
End If
End If
End Sub

Private Sub XPButton9_Click()
A = MsgBox("Apakah Akan Dihapus?", vbQuestion + vbYesNo, "Hapus Data")
If A = vbYes Then
Adodc2.Recordset.Delete
MsgBox "Data Telah Terhapus", vbInformation, "Informasi"
XPButton9.Enabled = False
Adodc1.RecordSource = "select* from transaksi_pembeli"
Adodc1.Refresh
Adodc2.RecordSource = "select* from transaksi_pembeli"
Adodc2.Refresh
End If
End Sub


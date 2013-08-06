VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AP MOTOR"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   8895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ENTRI DATA"
      TabPicture(0)   =   "Form1.frx":0442
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
      Tab(0).Control(9)=   "XPButton6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "XPButton1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "XPButton2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "XPButton3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "XPButton4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Adodc1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "XPButton5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "XPFrame1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "XPFrame4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "LIHAT DATA"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "XPFrame5"
      Tab(1).Control(1)=   "XPFrame3"
      Tab(1).Control(2)=   "XPButton12"
      Tab(1).Control(3)=   "XPButton11"
      Tab(1).Control(4)=   "XPButton10"
      Tab(1).Control(5)=   "XPButton9"
      Tab(1).Control(6)=   "XPButton8"
      Tab(1).Control(7)=   "XPButton7"
      Tab(1).Control(8)=   "DataGrid1"
      Tab(1).Control(9)=   "Adodc2"
      Tab(1).ControlCount=   10
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Masukkan nomor mesin  "
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Masukkan merk motor"
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Masukkan type motor"
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "Rp. "
         ToolTipText     =   "Masukkan harga motor"
         Top             =   3360
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":047A
         Left            =   1920
         List            =   "Form1.frx":0499
         TabIndex        =   4
         ToolTipText     =   "Masukkan warna motor dengan memilih"
         Top             =   2760
         Width           =   3135
      End
      Begin XPControls.XPFrame XPFrame4 
         Height          =   375
         Left            =   5400
         TabIndex        =   16
         Top             =   4920
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
            TabIndex        =   17
            Top             =   0
            Width           =   6015
         End
      End
      Begin XPControls.XPFrame XPFrame1 
         Height          =   1455
         Left            =   6240
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2566
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
         TabIndex        =   10
         ToolTipText     =   "Mengedit data"
         Top             =   4440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":04DD
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
         Left            =   7080
         Top             =   5640
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
         TabIndex        =   8
         ToolTipText     =   "Membersihkan tulisan pada teks isian"
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":352F
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
         TabIndex        =   7
         ToolTipText     =   "Mengupdate / memperbarui"
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form1.frx":6581
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
         TabIndex        =   9
         ToolTipText     =   "Menghapus data "
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form1.frx":95D3
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
         TabIndex        =   6
         ToolTipText     =   "Menyimpan data"
         Top             =   3960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":C625
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
         TabIndex        =   11
         ToolTipText     =   "Batal mengedit"
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":F677
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
      Begin XPControls.XPFrame XPFrame5 
         Height          =   495
         Left            =   -74760
         TabIndex        =   27
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
         Begin VB.Label Label14 
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
            Height          =   495
            Left            =   1320
            TabIndex        =   37
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label13 
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
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
      End
      Begin XPControls.XPFrame XPFrame3 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   29
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2566
         BackColor       =   14737632
         BackStyle       =   0
         Caption         =   "Pencarian Menggunakan  :"
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
      Begin XPControls.XPButton XPButton12 
         Height          =   975
         Left            =   -69960
         TabIndex        =   30
         ToolTipText     =   "Pencarian dengan menggunakan harga motor"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":126C9
         Caption         =   "Harga"
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
      Begin XPControls.XPButton XPButton11 
         Height          =   975
         Left            =   -71160
         TabIndex        =   31
         ToolTipText     =   "Pencarian dengan menggunakan type motor"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":1571B
         Caption         =   "Type"
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
      Begin XPControls.XPButton XPButton10 
         Height          =   975
         Left            =   -72360
         TabIndex        =   32
         ToolTipText     =   "Pencarian dengan menggunakan merk motor"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":1876D
         Caption         =   "Merk"
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
      Begin XPControls.XPButton XPButton9 
         Height          =   975
         Left            =   -72000
         TabIndex        =   33
         ToolTipText     =   "Menghapus data"
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Enabled         =   0   'False
         Picture         =   "Form1.frx":1B7BF
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
         Left            =   -74520
         TabIndex        =   34
         ToolTipText     =   "Pencarian dengan menggunakan nomor mesin"
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":1E811
         Caption         =   "Nomor Mesin"
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
      Begin XPControls.XPButton XPButton7 
         Height          =   975
         Left            =   -70560
         TabIndex        =   35
         ToolTipText     =   "Merefresh / menampilkan data kembali"
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Picture         =   "Form1.frx":21863
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
         Bindings        =   "Form1.frx":248B5
         Height          =   2775
         Left            =   -74760
         TabIndex        =   36
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
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Warna"
            Caption         =   "Warna"
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
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1800
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
         RecordSource    =   "select * from motor"
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Mesin"
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
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Merk"
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
         Left            =   480
         TabIndex        =   25
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   480
         TabIndex        =   24
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
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
         Left            =   480
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         Left            =   480
         TabIndex        =   22
         Top             =   3360
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
         Left            =   6600
         TabIndex        =   21
         Top             =   2040
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
         Left            =   6600
         TabIndex        =   20
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   1320
         Left            =   5280
         Picture         =   "Form1.frx":248CA
         Top             =   3480
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
         Left            =   1920
         TabIndex        =   19
         Top             =   600
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   960
      Top             =   7440
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1440
      Top             =   7440
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   1920
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
         TabIndex        =   12
         Top             =   0
         Width           =   8055
      End
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   0
      Picture         =   "Form1.frx":31F8C
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   9000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      X1              =   0
      X2              =   9000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTOR"
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
      TabIndex        =   14
      Top             =   0
      Width           =   3495
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
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "Form1.frx":38D75
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
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
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Adodc1.RecordSource = "Select * from Motor where No_Mesin='" & Text1 & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Informasi"
With Adodc1.Recordset
Text1 = !No_Mesin
Text2 = !Merk
Text3 = !Type
Combo1 = !Warna
Text4 = !Harga
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
Text4 = "Rp. " & ""
Text1.SetFocus
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
End If
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
Label14.Caption = Adodc2.Recordset.RecordCount
End Sub

Private Sub XPButton1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "Rp. " & "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
D = "select * from motor where No_Mesin='" & Text1 & "'"
Adodc1.RecordSource = D
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Tidak Bisa Disimpan, Karena Nomor Mesin Tersebut Sudah Ada", vbInformation, "Informasi"
Else
Adodc1.RecordSource = "select *from Motor where No_Mesin='" & Text1 & "'"
Adodc1.Refresh
With Adodc1.Recordset
.AddNew
!No_Mesin = Text1
!Merk = Text2
!Type = Text3
!Warna = Combo1
!Harga = Text4
.Update
End With
MsgBox "Data Tersimpan", vbInformation, "Informasi"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = "Rp. " & ""
Text1.SetFocus
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
Adodc2.Refresh
End If
End If
End Sub

Private Sub XPButton10_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Merk Motor Yang Akan Anda Cari!", "Pencarian")
D = "select * from motor where Merk='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
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

Private Sub XPButton11_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Type Motor Yang Akan Anda Cari!", "Pencarian")
D = "select * from motor where Type='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
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

Private Sub XPButton12_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Harga Motor Yang Akan Anda Cari Dengan Rp. Dahulu!", "Pencarian")
D = "select * from motor where Harga='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
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

Private Sub XPButton2_Click()
A = MsgBox("Apakah Akan Dihapus?", vbQuestion + vbYesNo, "Hapus Data")
If A = vbYes Then
Adodc1.Recordset.Delete
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = "Rp. " & ""
XPButton2.Enabled = False
XPButton3.Enabled = False
XPButton1.Enabled = True
XPButton4.Enabled = True
XPButton5.Visible = True
XPButton6.Visible = False
MsgBox "Data Telah Terhapus", vbInformation, "Information"
Text1.SetFocus
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
Adodc2.Refresh
End If
End Sub

Private Sub XPButton3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "Rp. " & "" Then
MsgBox "Data Tidak Lengkap", vbCritical, "Peringatan"
Else
Adodc1.RecordSource = "select *from Motor where No_Mesin='" & Text1 & "'"
If Adodc1.Recordset.RecordCount > 0 Then
With Adodc1.Recordset
!Merk = Text2
!Type = Text3
!Warna = Combo1
!Harga = Text4
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
Text4 = "Rp. " & ""
Text1.SetFocus
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
Adodc2.Refresh
End If
End If
End Sub

Private Sub XPButton4_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = "Rp. " & ""
Text1.SetFocus
End Sub

Private Sub XPButton5_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Or Text4 = "" Then
MsgBox "Jika Mau Edit, Masukkan Datanya Dengan Keypress Pada Text Nomor Mesin!", vbInformation, "Informasi"
 If KeyAscii = 13 Then
Adodc1.RecordSource = "Select * from Motor where No_Mesin='" & Text1 & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
MsgBox "Data Ada", vbInformation, "Peringatan"
With Adodc1.Recordset
Text1 = !No_Mesin
Text2 = !Merk
Text3 = !Type
Combo1 = !Warna
Text4 = !Harga
End With
XPButton2.Enabled = True
XPButton3.Enabled = True
XPButton1.Enabled = False
XPButton4.Enabled = False
XPButton5.Visible = False
XPButton6.Visible = True
Else
MsgBox "Data Tidak Ada", vbCritical, "Peringatan"
Text1 = ""
Text2 = ""
Text3 = ""
Combo1 = ""
Text4 = "Rp. " & ""
Text1.SetFocus
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
Text4 = "Rp. " & ""
Text1.SetFocus
End Sub

Private Sub XPButton7_Click()
Adodc2.RecordSource = "select* from motor"
Adodc2.Refresh
XPButton9.Enabled = False
End Sub

Private Sub XPButton8_Click()
XPButton9.Enabled = False
C = InputBox("Untuk Pencarian, Masukkan Nomor Mesin Yang Akan Anda Cari!", "Pencarian")
D = "select * from motor where No_Mesin='" & C & "'"
Adodc2.RecordSource = D
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
DataGrid1.Visible = True
Adodc2.Visible = True
E = MsgBox("Apakah Akan Diedit?", vbQuestion + vbYesNo, "Pertanyaan")
Else
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
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
Adodc1.RecordSource = "select* from motor"
Adodc1.Refresh
Adodc2.RecordSource = "select* from motor"
Adodc2.Refresh
End If
End Sub


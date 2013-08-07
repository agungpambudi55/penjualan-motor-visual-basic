VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form Form8 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form8.frx":0442
   ScaleHeight     =   3870
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   240
      Top             =   2760
   End
   Begin VB.Timer Timer3 
      Interval        =   70
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1560
   End
   Begin XPControls.ProgBarXP ProgBarXP1 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      ForeColor       =   0
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2011 AP Motor Corpration.All Rights Reserved"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   6615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "AP MOTOR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label1.Caption = "Start....."
End Sub

Private Sub Timer1_Timer()
ProgBarXP1.Value = ProgBarXP1.Value + 5
If ProgBarXP1.Value = 100 Then
Unload Me
Form6.Show
End If
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = "Start....." Then
Label1.Caption = "Loading....."
ElseIf Label1.Caption = "Loading....." Then
Label1.Caption = "Waiting....."
ElseIf Label1.Caption = "Waiting....." Then
Label1.Caption = "Load file....."
ElseIf Label1.Caption = "Load file....." Then
Label1.Caption = "Open Software....."
End If
End Sub

Private Sub Timer3_Timer()
If Shape1.Visible = True Then
Shape2.Visible = True
Shape3.Visible = True
Shape1.Visible = False
Shape12.Visible = False
Shape4.Visible = True
Shape11.Visible = False
ElseIf Shape2.Visible = True Then
Shape3.Visible = True
Shape4.Visible = True
Shape5.Visible = True
Shape2.Visible = False
ElseIf Shape3.Visible = True Then
Shape4.Visible = True
Shape5.Visible = True
Shape6.Visible = True
Shape3.Visible = False
ElseIf Shape4.Visible = True Then
Shape5.Visible = True
Shape6.Visible = True
Shape7.Visible = True
Shape4.Visible = False
ElseIf Shape5.Visible = True Then
Shape6.Visible = True
Shape7.Visible = True
Shape8.Visible = True
Shape5.Visible = False
ElseIf Shape6.Visible = True Then
Shape7.Visible = True
Shape8.Visible = True
Shape9.Visible = True
Shape6.Visible = False
ElseIf Shape7.Visible = True Then
Shape8.Visible = True
Shape9.Visible = True
Shape10.Visible = True
Shape7.Visible = False
ElseIf Shape8.Visible = True Then
Shape9.Visible = True
Shape10.Visible = True
Shape11.Visible = True
Shape8.Visible = False
ElseIf Shape9.Visible = True Then
Shape10.Visible = True
Shape11.Visible = True
Shape12.Visible = True
Shape9.Visible = False
ElseIf Shape10.Visible = True Then
Shape11.Visible = True
Shape12.Visible = True
Shape1.Visible = True
Shape10.Visible = False
ElseIf Shape11.Visible = True Then
Shape12.Visible = True
Shape1.Visible = True
Shape2.Visible = True
Shape11.Visible = False
ElseIf Shape12.Visible = True Then
Shape1.Visible = True
Shape2.Visible = True
Shape3.Visible = True
Shape12.Visible = False
End If
End Sub

Private Sub Timer4_Timer()
If Label2.Visible = True Then
Label2.Visible = False
ElseIf Label2.Visible = False Then
Label2.Visible = True
End If
End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1725
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4395
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1190.626
   ScaleMode       =   0  'User
   ScaleWidth      =   4127.132
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   510
      TabIndex        =   5
      Top             =   120
      Width           =   570
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   0
         Picture         =   "frmAbout.frx":014A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000040&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -549.345
      X2              =   4675.538
      Y1              =   828.262
      Y2              =   828.262
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "App Description:This program calculates the cost of internet usage as per the pulse."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3405
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Anaz Jaleel. Distribute only the compiled exe. I have not tested in NT and may not work in it."
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   0
      X1              =   -549.345
      X2              =   4661.452
      Y1              =   828.262
      Y2              =   828.262
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'From add form menu, I removed the sysinfo
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, _
                            Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub


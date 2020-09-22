VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   1395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Shape Shape2 
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   90
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   405
      Left            =   10
      Top             =   10
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Height          =   405
      Left            =   10
      TabIndex        =   0
      Top             =   10
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'So that nothing is left out which later haunts!

Private Sub Form_Load()
If frmNet.chpop.Value = 1 Then GoTo 10
Form1.Width = 1380
Form1.Height = 900 'So that only the required part of the form is shown even if
                    'we change size accidently during design

Form1.ScaleMode = 1 'Scale 1-Twip.Defining at runtime!You could do the same with
                    'properties box
Form1.WindowState = 0 '0-Normal window

Form1.Left = Screen.Width / 2 - 600 'Screen.Width - Form1.Width - 100
Form1.Top = Screen.Height 'we have to position the form above the tray

Form1.Height = 0 'Reset height to 0. We then have to increase it from 0 to max ht.
Form1.Visible = True 'Or else animation may not work

C_Ofrm Me, 1, False 'False to open and true to close form

'A simple FOR loop may too will do the trick
Timer1.Interval = 1000 'The time after which to unload the form
Timer1.Enabled = True
10 End Sub


Public Function C_Ofrm(frm As Form, Speed As Integer, tag As Boolean)
If Speed = 0 Then
    Exit Function 'The form will not be Closed/Opened
End If

If tag Then
    Do Until frm.Height <= 5
        DoEvents
        frm.Height = frm.Height - Speed * 1
        frm.Top = frm.Top + Speed * 1
    Loop
    Unload frm
Else
    Do Until frm.Height >= 900
        DoEvents
        frm.Height = frm.Height + Speed * 1
        frm.Top = frm.Top - Speed * 1
    Loop 'Mid$(frmNet.Label1.Caption, 8, Len(frmNet.Label1.Caption))
    Label1.Caption = frmNet.Label1.Caption
End If

End Function

Private Sub Timer1_Timer()
    C_Ofrm Me, 1, True 'Close or unload form
End Sub

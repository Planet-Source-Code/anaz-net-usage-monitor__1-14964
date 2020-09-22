VERSION 5.00
Begin VB.Form frmNet 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Net Usage Monitor"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   3150
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   29
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CheckBox chpop 
      BackColor       =   &H00C0C0FF&
      Caption         =   "No popup"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Always On top"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Update"
      Height          =   255
      Left            =   2160
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Load at startup with windows"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   2895
   End
   Begin VB.TextBox txtPd 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtPr 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   8400
   End
   Begin VB.TextBox txtC 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtTt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtTd 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtTc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H008080FF&
      Caption         =   "Move mouse over this to close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   8040
      Width           =   3165
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   28
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reset log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1920
      TabIndex        =   27
      Top             =   4440
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pop up"
      Height          =   240
      Left            =   1080
      TabIndex        =   25
      Top             =   5160
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      Height          =   240
      Left            =   1800
      TabIndex        =   24
      Top             =   5160
      Width           =   390
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   8
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   840
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
      Height          =   240
      Left            =   2445
      TabIndex        =   21
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      Height          =   240
      Left            =   2445
      TabIndex        =   20
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
      Height          =   240
      Left            =   2450
      TabIndex        =   19
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "hr:min:s"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2450
      TabIndex        =   18
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " About"
      Height          =   240
      Left            =   2280
      TabIndex        =   17
      Top             =   5160
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status : "
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
      Height          =   240
      Left            =   1560
      TabIndex        =   13
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblPd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duration of Pulse"
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   2880
      Width           =   1485
   End
   Begin VB.Label LblPr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate/Pulse"
      Height          =   240
      Left            =   0
      TabIndex        =   11
      Top             =   3360
      Width           =   930
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   15
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label lblTt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label lblTd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disconected"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblTc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connected"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   870
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   3150
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   2175
      Left            =   8
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Top             =   5400
      Width           =   3165
   End
   Begin VB.Menu popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu hide 
         Caption         =   "Hide"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Only for you.Please don't distribute this code as such
'Only exe be distributed and don't take credit for parts that
'u didn't develop and please don't change the original log file details.
'You can add your own stuff there.
'Any modification suggestions are always welcome.
'If you make modification send the new code to me so that I too can study
'something new.

Dim Pic As Boolean  'To twinkle the two computers

Private Sub Form_Load()
Me.WindowState = 0                              'Normal window state

Set Iconobj = LoadResPicture(101, 1)         'Load from res file
Image1.Picture = LoadResPicture(101, 1)
AddIcon frmNet, Iconobj.Handle, Iconobj   'Add to tray
                                                'Set position to the topmost.
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Width / 15, _
                            Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW

If Dir("c:\IULog.txt") = "" Then 'It is the first time u are running this app
    Open "c:\IULog.txt" For Append As #1 'I think that c drive is there in every
                                         'machine.Use app.path if u think not :-)
'    A little bit about me. If you are not a very busy person,
'    don 't forget to give me a postcard of your place(or a stamp or a coin.)
'    or a photo of your country side.  It's my hobby to collect stamps, coins
'    and photos of different place/countries. The email killed the snailmail
'    but I love to have them going so that I am not bored by 'outlooks'.
'    I do hope that you have already given 5 globes for my effort.
'    Thank you for that. Please vote and encourage me to submit
'    further beautiful codes.
'    Visit http://ajsoftware.freeserver.com/index. It contains an
'    activex counter,developed by me.
        
    Print #1, "===================================" & vbCrLf & _
              "Log Of Internet Usage. Ver 1.00" & vbCrLf & _
              "Designed and created by Anaz Jaleel," & vbCrLf & _
              "Anazview, Punnathala, Kollam,Kerala," & vbCrLf & _
              "India.Pin-691012.If you get time," & vbCrLf & _
              "send me a postcard or a photo of " & vbCrLf & _
              "your place or drop a mail at" & vbCrLf & _
              "anaz@operamail.com" & vbCrLf & vbCrLf & _
              "Please notify any bugs." & vbCrLf & _
              "Logging Started on" & vbCrLf & "   " & _
              Format$(Now, "dddd, mmm d, yyyy") & vbCrLf & _
              "===================================" & vbCrLf
    Close #1
    'Some thing are to be remembered and this is better that usual
    'savesetting and getsetting. Here we are actually reseting the log
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Amount", 0
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "TotalTime", 0
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "Duration", "180"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "PulseRate", "1.20"
    

End If
If GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "NetUsage", "") = "" Then
Check1.Value = 0        'Dont show at startup
Else
Check1.Value = 1        'Show at startup is selected previously and now we are
                        'running with windows
End If

                        'Fill the form details with previous values
txtPd.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "Duration", "180")
txtPr.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "PulseRate", "1.20")
txtTt.Text = ConvertTime(GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "TotalTime", "0"))
txtC.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Amount", "0")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Connected Then
                        'The connection is not terminated,but our program is!
    LogStopTime         'Any way record it
    Open "c:\IULog.txt" For Append As #1 'Append to the file the details of this
                                         'Session, together with this note
    Print #1, "The program may have been terminated " & vbCrLf & _
              "before internet is disconnected." & vbCrLf & _
              "Log may be incorrect!" & vbCrLf & _
              "ºººººººººººººººººººººººººººººººººººººººººº"
    Close #1
End If
End Sub

Private Sub hide_Click()
Me.hide                 'Hide form
End Sub

Private Sub Label11_Click()
Me.hide                 'Hide form
End Sub

Private Sub Label12_Click()
Dim choice As VbMsgBoxResult
If Connected Then GoTo 10   'Foul play
                            'Can use password protection
choice = MsgBox("Reset Log button is clicked.Are you sure?", vbYesNo, "Reseting..........")
If choice = vbYes Then ReSetlog
10 End Sub


Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.Enabled = False
Dim i As Integer
For i = 8610 To 5850 Step -1
DoEvents
Me.Height = i
Next i
Text1.Text = ""
Label4.Enabled = True 'To enable the next part to be closed
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 3020 To 5850
DoEvents                 'Slowly drop down the needed part
Me.Height = i
Next i
Label4.Enabled = True   'Enable the closing 'key'
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Enabled = False  'Or else you will get another effect!Try it.

Dim i As Integer
For i = 5850 To 3020 Step -1 'To go backwards
DoEvents                'Fold up the form.For this effect you can use Do Until
Me.Height = i           'See the following for such an example.
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=14911
'This tries to copy the MSN Messenger's way of showing form when you log on with
'an hyperlink int the form. If u use a little brain then you can design a code
'ticker as seen on this place
'Rate that too.
'I found it useful here also.See later parts of the code
Next i
End Sub

Private Sub Label5_Click()

Load Form1
Form1.Timer1.Enabled = True
End Sub

Private Sub Label6_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
Connected = IsConnected
l = GetCursorPos(a)
If (a.X > (Screen.Width * 0.9 / Screen.TwipsPerPixelX)) And _
(a.Y > (Screen.Height * 0.97 / Screen.TwipsPerPixelY)) Then
'SetCursorPos (Screen.Width * 0.9 / Screen.TwipsPerPixelX) - 5, (Screen.Height * 0.97 / Screen.TwipsPerPixelY) - 5
'The above code will throw the cursor around.It annoy people
Load Form1                      'A simple pop up message
Form1.Timer1.Enabled = True
End If
'Connected = CBool(Check3.Value) 'I used this to test it over and over again
                                'So that it is bug free.Notify if any found
                                'Note:I work on Win98 machine
If Connected Then
    Label1.Caption = "Online :" & ConvertTime(Timer - TimeConnected)
                            'Difference in time to update the label
    If Pic Then
        Image1.Picture = LoadResPicture(102, 1): Pic = False
                            'Toggle with the pic.
                            'I first inteded to modify the tray also
                            'but if you are careless and internet connection
                            'is slow surely there will be confusion
    Else
        Image1.Picture = LoadResPicture(103, 1): Pic = True
    End If
End If
If Connected And LogStarted = False Then
    TimeConnected = Timer   'This is another method to find the duration
                            'without any api call.
                            'The time we connected is stored here
                            'and any time we need to find duration
                            'Timer-this value will get you that
    Image1.Picture = LoadResPicture(101, 1)
    frmNet.txtTd.Text = ""  'Clear the disconnected time field
    Label1.Caption = "Online" 'You are online.
    frmNet.Show             'In case this is hidden to notify that
                            'log has started
    LogStarted = True       'So that these lines are excluded next time
    LogStarttime            'Note the necessary details
ElseIf Connected = False And LogStarted Then
    LogStopTime             'Disconnected and hence close log
    LogStarted = False      'Restore everything as not connected
    Image1.Picture = LoadResPicture(101, 1)
    Label1.Caption = "Offline"
    'frmNet.Hide            'If you want to hide the form
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
delIcon Iconobj.Handle 'Clean the mess as long as it is possible

Unload Form_main
Unload frmNet
Unload frmAbout
Unload Form1

Set Form_main = Nothing
Set frmNet = Nothing
Set frmAbout = Nothing
Set Form1 = Nothing
End

End Sub
Private Sub exit_Click()
Form_Unload (0)             'Terminate app
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
      Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP:
        Me.PopupMenu popup  'To display the option menu
    Case WM_LBUTTONDBLCLK:
         frmNet.Show        'If dblclicked ,show form without wasting any time.
    End Select
End Sub

Private Sub open_Click()
frmNet.Show                 'Show the form
End Sub

Private Sub Check1_Click()
'To run at startup or not
If Check1.Value = 1 Then
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "NetUsage", App.Path & "\" & App.EXEName & ".exe"
Else
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "NetUsage", ""
End If

End Sub

Private Sub Check2_Click()
'Top or not
If Check2.Value = 1 Then
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / 15, _
                            Me.Top / 15, Me.Width / 15, _
                            Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
SetWindowPos Me.hWnd, HWND_NOTOPMOST, Me.Left / 15, _
                            Me.Top / 15, Me.Width / 15, _
                            Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End Sub

Private Sub Command1_Click()
'Save the pulse time and cost per pulse
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "Duration", txtPd.Text
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "PulseRate", txtPr.Text
End Sub
Private Sub ReSetlog()
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Amount", 0
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "TotalTime", 0
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "Duration", "180"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Cost", "PulseRate", "1.20"
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Session", "No", 0
    txtTt.Text = "00:00:00"
    txtC.Text = "0"
    Open "c:\IULog.txt" For Append As #1
    Print #1, "===================================" & vbCrLf & _
              "The values are reset on:" & vbCrLf & _
              Format$(Now, "dddd, mmm d, yyyy") & vbCrLf & _
              "===================================" & vbCrLf
    Close #1
    
End Sub
Private Sub Label13_Click()
Dim temp As String
Dim i As Integer
Open "c:\iulog.txt" For Input As #1
While Not EOF(1)
    Line Input #1, temp
    Text1.SelText = temp & vbCrLf
Wend
Close #1
For i = 5850 To 8610
DoEvents                 'Slowly drop down the needed part
Me.Height = i
Next i
Label14.Enabled = True
Label4.Enabled = False 'Or else the form close accidently
End Sub


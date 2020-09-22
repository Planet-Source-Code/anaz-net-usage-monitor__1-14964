Attribute VB_Name = "ModCheck"
Option Explicit
Public Type POINT
    x As Long
    y As Long
End Type

Public a As POINT
Public l As Long

Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'DWORD APIENTRY RasEnumConnectionsA( LPRASCONNA, LPDWORD, LPDWORD );
'#define LPRASCONNA RASCONNA*
'RASCONNA
'{
'    DWORD    dwSize;
'    HRASCONN hrasconn;
'    CHAR     szEntryName[ RAS_MaxEntryName + 1 ];
'
'#if (WINVER >= 0x400)
'    CHAR     szDeviceType[ RAS_MaxDeviceType + 1 ];
'    CHAR     szDeviceName[ RAS_MaxDeviceName + 1 ];
'#End If
'#if (WINVER >= 0x401)
'    CHAR     szPhonebook [ MAX_PATH ];
'    DWORD    dwSubEntry;
'#End If
'};

'DWORD APIENTRY RasGetConnectStatusA( HRASCONN, LPRASCONNSTATUSA );
'O brother/sister! no time here. Refer ras.h in VC++ dirs for details and
'make it work for NT too
Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long

Public start As Double      'To store the start time(Here we use GetTickCount)
Public total As Double      'Other things you need
Public TtlSeconds As Double
Public totalsec As String
Public Session As Integer
Public Connected As Boolean
Public LogStarted As Boolean
Public actual As String
Public sessioncost As Double
Public TimeConnected As Date

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'#define RAS_MaxEntryName      256
'#define RAS_MaxDeviceType     16
'#define RAS_MaxDeviceName     32
Public Const RAS_MaxEntryName = 256
Public Const RAS_MaxDeviceType = 16
Public Const RAS_MaxDeviceName = 32

Public Type RASCONN
   dwSize As Long
   hRasCon As Long
   szEntryName(1 To RAS_MaxEntryName) As Byte ' ANSI entry point .'. use Bytes
   szDeviceType(1 To RAS_MaxDeviceType) As Byte
   szDeviceName(1 To RAS_MaxDeviceName) As Byte
End Type

Public Type RASCONNSTATUS
   dwSize As Long
   RasConnState As Long
   dwError As Long
   szDeviceType(1 To RAS_MaxDeviceType) As Byte
   szDeviceName(1 To RAS_MaxDeviceName) As Byte
End Type



Public Function IsConnected() As Boolean
'This is the part where we check for live connection
   Dim TRasCon(255) As RASCONN
   Dim lg As Long
   Dim lpcon As Long
   Dim RetCDec As Long
   Dim Tstatus As RASCONNSTATUS

   TRasCon(0).dwSize = 412
   lg = 256 * TRasCon(0).dwSize

   RetCDec = RasEnumConnections(TRasCon(0), lg, lpcon)
   Tstatus.dwSize = 160
   RetCDec = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)

   If Tstatus.RasConnState = &H2000 Then
      IsConnected = True
   Else
      IsConnected = False
   End If

End Function

Public Sub LogStarttime()   'Open and write the log start details

frmNet.txtTc = Time         'The connected time is to be displayed
totalsec = 0                'Initialise totalsec and get Session detail
Session = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Session", "No", 0)
start = GetTickCount&       'Note the value
Session = Session + 1       'This is the next session
Open "c:\IULog.txt" For Append As #1    'Write all details in the file
'A little art makink it easier to see details in log ;-)
Print #1, "___________________________________"
Print #1, "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
Print #1, "Session     : " & Session
Print #1, "Date        : "; Format$(Now, "ddd, mmm d, yyyy")
Print #1, "Connected   : " & Time
Close #1
'Save to be sure;actually not needed
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Start", CStr(start)

End Sub

Public Sub LogStopTime()    'Open and write the log stop details

Dim finish As Double        'For all that local calculations
Dim tempSec As Double
Dim temp As Integer
Dim todaytime As Double
Dim totalcost As Double

frmNet.txtTd = Time         'The connected time is to be displayed
'Get the save value of ticks
start = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Start", 1)
finish = GetTickCount&      'Hey Window, What is the tic now ?
total = (finish - start)    'This is what we need
totalsec = Round(total / 1000, 0) 'We dont need any decimal
actual = totalsec           'We need the actual value to write in log file
                            'We are going to find total time as per pulse
                            'Get the total time till previous session
TtlSeconds = GetSettingString(HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "TotalTime", 0)
temp = totalsec Mod CInt(frmNet.txtPd.Text) 'Find the extra time after previous
                            'pulse.If its is 0 proceed else make it as per pulse
If temp <> 0 Then todaytime = totalsec: totalsec = totalsec + CInt(frmNet.txtPd.Text) - temp

tempSec = totalsec + TtlSeconds
                            'Finally the total usage time is here

TtlSeconds = tempSec        'These are not needed if the pulse is constant
                            'But in any case we changed pulse this will update
                            'our change.Same procedure as above
temp = TtlSeconds Mod CInt(frmNet.txtPd.Text)
If temp <> 0 Then TtlSeconds = TtlSeconds + CInt(frmNet.txtPd.Text) - temp
                            'Calculate the session cost
sessioncost = totalsec * CDbl(frmNet.txtPr.Text) / CDbl(frmNet.txtPd.Text)
Open "c:\IULog.txt" For Append As #1
Print #1, "Disconnected: " & Time
Print #1, "___________________________________"
Print #1, "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" & vbCrLf & _
        "Actual Time this session: " & vbCrLf & _
        "hr:min: sec" & vbCrLf & _
        ConvertTime(todaytime) & _
        vbCrLf & _
        "Pulse time this session: " & vbCrLf & _
        "hr:min:sec" & vbCrLf & _
        ConvertTime(totalsec) & _
        vbCrLf & _
        "Session Cost: " & sessioncost & vbCrLf
Close #1
'This calculates the total cost and prints it
totalcost = TtlSeconds * CDbl(frmNet.txtPr.Text) / CDbl(frmNet.txtPd.Text)
Open "c:\IULog.txt" For Append As #1
Print #1, "Total Pulse time connected " & vbCrLf & _
        "since log started:" & vbCrLf & _
        "hr : min : sec" & vbCrLf & ConvertTime(TtlSeconds) & vbCrLf & _
        "Total Cost :" & totalcost & vbCrLf & _
        "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Close #1
frmNet.txtC.Text = totalcost 'Display the total cost
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "Amount", CStr(totalcost)
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Session", "No", CStr(Session)
SaveSettingString HKEY_LOCAL_MACHINE, "Software\AjSoft\NetUsage\Data\Time", "TotalTime", CStr(TtlSeconds)
frmNet.txtTt = ConvertTime(TtlSeconds)
DrawingText
End Sub
Public Function ConvertTime(ByVal tempSec As Single) As String
'This converts the time to hr:min:sec format
Dim h As String
Dim m As String
Dim s As String
On Error GoTo ErrHandle
h = CStr(Int(tempSec / 3600))
tempSec = tempSec - h * 3600
m = CStr(Int(tempSec / 60))
tempSec = tempSec - m * 60
s = CStr(Int(tempSec))
If Len(h) = 1 Then h = "0" & h  'So that a time 0:0:0 will be converted to
If Len(m) = 1 Then m = "0" & m  '00:00:00 the actual intended way
If Len(s) = 1 Then s = "0" & s
ConvertTime = h & ":" & m & ":" & s
ErrHandle: 'May we not reach here.Actually no need. Remove it
End Function




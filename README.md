<div align="center">

## Net Usage Monitor


</div>

### Description

Net Usage Moniter-ReadMe

This application calculates the time connected to the net and finds out the total amount spend by surfing the net. The cost is calculated as per the pulse. If 180 is the pulse then even if you connect for 1 second the cost of 180 second is added.Thus a correct amout spend is found out.

It also uses DD7 to display the session details after you disconnect from the net and also a log file.

When you move the mouse near the tray the time you are connected to the net is poped up. You can also disable this popup.

The code is well commented-each and every line. You can learn the following things from this code.

*Check for internet connection

*Read and write from registry

*Keeping log files

*Display text directly to the screen

*Add details of a file to textbox

*Add, Delete,Modify the tray icon

*Read from resource file

*Slide down opening of form as given by Scrolling Notice http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=14911

The program is not tested on NT and is not for it. I think changes are needed in api call,then you have to get OSversion etc, to make it work there and here in 98. I created this for my personal use and me and my friends dont have NT to test the other part.

anaz@operamail.com
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-02-03 15:53:10
**By**             |[Anaz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/anaz.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD14473232001\.zip](https://github.com/Planet-Source-Code/anaz-net-usage-monitor__1-14964/archive/master.zip)

### API Declarations

```
Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Long
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
```






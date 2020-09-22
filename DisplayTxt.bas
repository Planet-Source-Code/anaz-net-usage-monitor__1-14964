Attribute VB_Name = "Text"
Option Explicit
'A little trick with the display.Might end up with 'beautiful' screens just as
'our,O sorry, my room!. Press refresh to get the old good screen
'Needs directX7.We all play game :-).You have it in you!(your system,I mean)

Public Type DxObj       'Everything in one
    TextFont As IFont
    DText As DirectX7
    DDraw As DirectDraw7
    MainSurface As DirectDrawSurface7
    SurfaceDesc As DDSURFACEDESC2
End Type

Public DText As DxObj   'DirectX objects

Public Sub DrawingText()

Dim i As Integer
Dim X As Integer 'Locals to see the dimesion of the screen
Dim Y As Integer
Dim V As Integer
Dim W As Integer
'To get the no. of pixels
X = Screen.Width / Screen.TwipsPerPixelX
Y = Screen.Height / Screen.TwipsPerPixelY

Set DText.TextFont = New StdFont 'Create the font used to draw text
Set DText.DText = New DirectX7 'Create main directX -object
Set DText.DDraw = DText.DText.DirectDrawCreate("") 'Create main DirectDraw -object

'Set co-operative level
DText.DDraw.SetCooperativeLevel Form_main.hWnd, DDSCL_NORMAL

'Create target surface, where to blit the text.Here we select the screen.
With DText.SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With
Set DText.MainSurface = DText.DDraw.CreateSurface(DText.SurfaceDesc)

With DText.TextFont
    .Size = 15
    .Name = "Monotype Corsiva"
    .Bold = False
    .Underline = False
End With

DText.MainSurface.SetFont DText.TextFont

V = X * 0.4 'Place to display the text. A little offset to the center of the
W = Y * 0.4 'screen
'This is to blink the string
For i = 0 To 4
    DrawText "Disconnected!", V, W, RGB(255, 255, 255), RGB(0, 0, 0) 'Draw  text
    Sleep 100
    DrawText "Disconnected!", V, W, RGB(0, 0, 0), RGB(255, 255, 255) 'Invert text to blink it
    Sleep 100
Next
    DrawText "Actual Used Time       = " & actual, V, W + 20, RGB(255, 255, 255), RGB(0, 0, 0) '
    Sleep 100
    DrawText "Used Time as per pulse = " & totalsec, V, W + 40, RGB(255, 255, 255), RGB(0, 0, 0) 'Draw  text
    Sleep 100
    DrawText "Cost as per pulse time = " & CStr(sessioncost), V, W + 60, RGB(255, 255, 255), RGB(0, 0, 0) 'Draw  text

With DText.TextFont
    .Size = 11
    .Name = "Monotype Corsiva"
    .Bold = False
    .Underline = True
    .Weight = 100
End With

DText.MainSurface.SetFont DText.TextFont
'To draw normally so that the text of small size is readable
DrawText1d "Designed by Anaz Jaleel.Refer program log for details.", V, W + 85, RGB(255, 0, 255)
End Sub

Sub DrawText(Text As String, X As Integer, Y As Integer, outerColor As Long, innerColor As Long)
'Off setting the text by some value so that an inner and outer effect is produced
DText.MainSurface.SetForeColor outerColor
DText.MainSurface.DrawText X, Y, Text, False
DText.MainSurface.DrawText X + 2, Y, Text, False
DText.MainSurface.DrawText X, Y + 2, Text, False
DText.MainSurface.DrawText X + 2, Y + 2, Text, False

DText.MainSurface.SetForeColor innerColor
DText.MainSurface.DrawText X + 1, Y + 1, Text, False
End Sub

Sub DrawText1d(Text As String, X As Integer, Y As Integer, Color As Long)

DText.MainSurface.SetForeColor Color
DText.MainSurface.DrawText X, Y, Text, False
 End Sub









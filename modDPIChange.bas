Attribute VB_Name = "modDPIChange"
Option Explicit

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Private Const CCHFORMNAME = 32
'Private Const CCHDEVICENAME = 32
'Private Const MONITOR_CCHDEVICENAME As Long = 32    ' device name fixed length
'Private Const ENUM_CURRENT_SETTINGS As Integer = -1
'Private Const WM_MOVE As Long = &H3, WM_MOVING As Long = &H216, WM_DISPLAYCHANGE As Long = &H7E
''Private Const WM_GETMINMAXINFO As Long = &H24
''Private Const GWL_WNDPROC As Long = -4
''Private Const MIN_WIDTH As Long = 13500 ' Minimum width in twips (15 twips = 1 pixel)
'
'Private Type POINT
'    X As Long
'    Y As Long
'End Type

'Private Type MONITORINFOEX
'    cbSize As Long
'    rcMonitor As RECT
'    rcWork As RECT
'    dwFlags As Long
'    szDevice As String * MONITOR_CCHDEVICENAME
'End Type
'
'Private Type DEVMODEW_FORDISP
'    dmDeviceName(CCHDEVICENAME - 1) As Integer
'    dmSpecVersion As Integer
'    dmDriverVersion As Integer
'    dmSize As Integer
'    dmDriverExtra As Integer
'    dmFields As DevModeFields
'    'PrinterOrDisplayFields(0 To 15) As BYTE 'union {
'                                            '   ' printer only fields */
'                                            '   struct {
'                                            '     short dmOrientation;
'                                            '     short dmPaperSize;
'                                            '     short dmPaperLength;
'                                            '     short dmPaperWidth;
'                                            '     short dmScale;
'                                            '     short dmCopies;
'                                            '     short dmDefaultSource;
'                                            '     short dmPrintQuality;
'                                            '   } DUMMYSTRUCTNAME;
'                                            '   ' display only fields */
'                                            '   struct {
'    dmPosition As POINT                     '     POINTL dmPosition;
'    dmDisplayOrientation As DevModeDisplayOrientation '    DWORD  dmDisplayOrientation;
'    dmDisplayFixedOutput As Long            '     DWORD  dmDisplayFixedOutput;
'                                            '   } DUMMYSTRUCTNAME2;
'                                            ' } DUMMYUNIONNAME;
'    dmColor As Integer
'    dmDuplex As Integer
'    dmYResolution As Integer
'    dmTTOption As Integer
'    dmCollate As Integer
'    dmFormName(CCHFORMNAME - 1) As Integer
'    dmLogPixels As Integer
'    dmBitsPerPel As Long
'    dmPelsWidth As Long
'    dmPelsHeight As Long
'    DisplayFlagsOrNup As Long   ' union {
'                                '     DWORD  dmDisplayFlags;
'                                '     DWORD  dmNup;
'                                ' } DUMMYUNIONNAME2;
'    dmDisplayFrequency As Long
'    '#if(WINVER >= 0x0400)
'    dmICMMethod As DevModeICMMethods
'    dmICMIntent As DevModeICMIntents
'    dmMediaType As DevModePrintMediaTypes
'    dmDitherType As DevModePrintDitherTypes
'    dmReserved1 As Long
'    dmReserved2 As Long
'    '#if (WINVER >= 0x0500) || (_WIN32_WINNT >= _WIN32_WINNT_NT4)
'    dmPanningWidth As Long
'    dmPanningHeight As Long
'    '#endif
'    '#endif /* WINVER >= 0x0400 */
'End Type
'
'
'Private Enum DevModeFields
'    DM_ORIENTATION = &H1
'    DM_PAPERSIZE = &H2
'    DM_PAPERLENGTH = &H4
'    DM_PAPERWIDTH = &H8
'    DM_SCALE = &H10
'    DM_POSITION = &H20
'    DM_NUP = &H40
'    DM_DISPLAYORIENTATION = &H80
'    DM_COPIES = &H100
'    DM_DEFAULTSOURCE = &H200
'    DM_PRINTQUALITY = &H400
'    DM_COLOR = &H800
'    DM_DUPLEX = &H1000
'    DM_YRESOLUTION = &H2000
'    DM_TTOPTION = &H4000
'    DM_COLLATE = &H8000&
'    DM_FORMNAME = &H10000
'    DM_LOGPIXELS = &H20000
'    DM_BITSPERPEL = &H40000
'    DM_PELSWIDTH = &H80000
'    DM_PELSHEIGHT = &H100000
'    DM_DISPLAYFLAGS = &H200000
'    DM_DISPLAYFREQUENCY = &H400000
'    DM_ICMMETHOD = &H800000
'    DM_ICMINTENT = &H1000000
'    DM_MEDIATYPE = &H2000000
'    DM_DITHERTYPE = &H4000000
'    DM_PANNINGWIDTH = &H8000000
'    DM_PANNINGHEIGHT = &H10000000
'    DM_DISPLAYFIXEDOUTPUT = &H20000000
'End Enum
'
'Private Enum DevModeDisplayOrientation
'    DMDO_0 = 0
'    DMDO_DEFAULT = DMDO_0
'    DMDO_90 = 1
'    DMDO_180 = 2
'    DMDO_270 = 3
'End Enum
'Private Enum DevModeICMMethods
'    DMICMMETHOD_NONE = 1 '/* ICM disabled */
'    DMICMMETHOD_SYSTEM = 2 '/* ICM handled by system */
'    DMICMMETHOD_DRIVER = 3 '/* ICM handled by driver */
'    DMICMMETHOD_DEVICE = 4 '/* ICM handled by device */
'    DMICMMETHOD_USER = 256 '/* Device-specific methods start here */
'End Enum
'Private Enum DevModeICMIntents
'    DMICM_SATURATE = 1 '/* Maximize color saturation */
'    DMICM_CONTRAST = 2 '/* Maximize color contrast */
'    DMICM_COLORIMETRIC = 3 '/* Use specific color metric */
'    DMICM_ABS_COLORIMETRIC = 4 '/* Use specific color metric */
'    DMICM_USER = 256 '/* Device-specific intents start here */
'End Enum
'Private Enum DevModePrintMediaTypes
'    DMMEDIA_STANDARD = 1 '/* Standard paper */
'    DMMEDIA_TRANSPARENCY = 2 '/* Transparency */
'    DMMEDIA_GLOSSY = 3 '/* Glossy paper */
'    DMMEDIA_USER = 256 '/* Device-specific media start here */
'End Enum
'Private Enum DevModePrintDitherTypes
'    DMDITHER_NONE = 1 '/* No dithering */
'    DMDITHER_COARSE = 2 '/* Dither with a coarse brush */
'    DMDITHER_FINE = 3 '/* Dither with a fine brush */
'    DMDITHER_LINEART = 4 '/* LineArt dithering */
'    DMDITHER_ERRORDIFFUSION = 5 '/* LineArt dithering */
'    DMDITHER_RESERVED6 = 6 '/* LineArt dithering */
'    DMDITHER_RESERVED7 = 7 '/* LineArt dithering */
'    DMDITHER_RESERVED8 = 8 '/* LineArt dithering */
'    DMDITHER_RESERVED9 = 9 '/* LineArt dithering */
'    DMDITHER_GRAYSCALE = 10 '/* Device does grayscaling */
'    DMDITHER_USER = 256 '/* Device-specific dithers start here */
'End Enum
''Private Type POINTAPI
''    x As Long
''    y As Long
''End Type
''
''Private Type MINMAXINFO
''    ptReserved As POINTAPI
''    ptMaxSize As POINTAPI
''    ptMaxPosition As POINTAPI
''    ptMinTrackSize As POINTAPI
''    ptMaxTrackSize As POINTAPI
''End Type
''
''Private lpPrevWndProc As Long
'
'Private Declare Function GetScaleFactorForMonitor Lib "shcore.dll" (ByVal hMonitor As Long, ByRef pScale As Long) As Long
'Private Declare Function GetModuleHandleW Lib "kernel32.dll" (ByVal lpModuleName As Long) As Long
'Private Declare Function LoadLibraryExW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
'Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
'    ByVal lpPrevWndProc As Long, ByVal hWnd As Long, _
'    ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'uncomment below to re-enable subclassing
'Private Const WM_NCDESTROY As Long = &H82, WM_DPICHANGED As Long = &H2E0
'Private Const SWP_NOACTIVATE As Long = &H10, SWP_NOOWNERZORDER As Long = &H200, SWP_NOZORDER As Long = &H4

'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
'Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
'Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private lNewDPI As Long, xControl As Control
'Public sngScaleFactor As Single

'Public Function DPI_SubclassForm(objForm As Form, Optional dwRefData As Long) As Boolean
'On Error GoTo error:
'If dwRefData = 0 Then dwRefData = GetDeviceCaps(objForm.hDC, LOGPIXELSX)
'DPI_SubclassForm = SetWindowSubclass(objForm.hWnd, AddressOf DPI_WndProc, ObjPtr(objForm), dwRefData)
'out:
'On Error Resume Next
'Exit Function
'error:
'Call HandleError("DPI_SubclassForm")
'Resume out:
'End Function
'
'Public Function DPI_UnSubclassForm(hWnd As Long, uIdSubclass As Long) As Boolean
'On Error Resume Next
'DPI_UnSubclassForm = RemoveWindowSubclass(hWnd, AddressOf DPI_WndProc, uIdSubclass)
'End Function
'
'Private Function DPI_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As RECT, ByVal objForm As Form, ByVal dwRefData As Long) As Long
'    On Error GoTo error:
'    Dim rNew As RECT
'    Select Case uMsg
'        Case WM_NCDESTROY ' Remove subclassing as the window is about to be destroyed
'            DPI_UnSubclassForm hWnd, ObjPtr(objForm)
'        Case WM_DPICHANGED ' This message signals a change in the DPI of the current monitor or the window was dragged to a monitor with a different DPI
''            lNewDPI = wParam And &HFFFF&: sngScaleFactor = lNewDPI / dwRefData: DPI_SubclassForm objForm, lNewDPI ' Calculate the new DPI and ScaleFactor and update dwRefData
''            With lParam
''                SetWindowPos hWnd, 0, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER ' Resize the form to reflect the new DPI changes
''            End With
''            objForm.Font.Size = objForm.Font.Size * sngScaleFactor
''            For Each xControl In objForm.Controls ' After the form is resized do the same for all its controls
''                If Not (TypeOf xControl Is Timer Or TypeOf xControl Is Menu) Then  ' Do not process controls without dimensions
''                    With xControl
''                        .Left = .Left * sngScaleFactor: .Top = .Top * sngScaleFactor: .Width = .Width * sngScaleFactor ' Common properties for most controls
''                        Select Case True
''                            Case TypeOf xControl Is Label, TypeOf xControl Is TextBox, TypeOf xControl Is CommandButton, _
''                                 TypeOf xControl Is Frame, TypeOf xControl Is PictureBox, _
''                                 TypeOf xControl Is ListBox, TypeOf xControl Is OptionButton, TypeOf xControl Is CheckBox
''                                .Font.Size = .Font.Size * sngScaleFactor: .Height = .Height * sngScaleFactor
''                            Case TypeOf xControl Is ComboBox ' Height is ReadOnly for a ComboBox
''                                .Font.Size = .Font.Size * sngScaleFactor
''                            Case TypeOf xControl Is HScrollBar, TypeOf xControl Is Image ' These controls don't have a Font property
''                                .Height = .Height * sngScaleFactor
''                        End Select
''                    End With
''                End If
''            Next xControl
'    End Select
'    DPI_WndProc = DefSubclassProc(hWnd, uMsg, wParam, VarPtr(lParam))
'out:
'    On Error Resume Next
'    Exit Function
'error:
'    Call HandleError("DPI_WndProc")
'    Resume out:
'End Function

Public Function ConvertScale(sngValue As Single, ScaleFrom As ScaleModeConstants, ScaleTo As ScaleModeConstants) As Single
On Error GoTo error:
Dim hDC As Long, lDPI_X As Long, lDPI_Y As Long, sngTPP_X As Single, sngTPP_Y As Single
Const HimetricPerPixel As Single = 26.45834

hDC = GetDC(0)
lDPI_X = GetDeviceCaps(hDC, LOGPIXELSX): lDPI_Y = GetDeviceCaps(hDC, LOGPIXELSY)
sngTPP_X = 1440 / lDPI_X: sngTPP_Y = 1440 / lDPI_Y
hDC = ReleaseDC(0, hDC)

Select Case True
    Case ScaleFrom = ScaleTo
        ConvertScale = sngValue
    Case (ScaleFrom = vbTwips) And (ScaleTo = vbPixels)
        ConvertScale = sngValue / sngTPP_X
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbTwips)
        ConvertScale = sngValue * sngTPP_X
    Case (ScaleFrom = vbTwips) And (ScaleTo = vbPoints)
        ConvertScale = sngValue / 20
    Case (ScaleFrom = vbPoints) And (ScaleTo = vbTwips)
        ConvertScale = sngValue * 20
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbPoints)
        ConvertScale = sngValue * sngTPP_X / 20
    Case (ScaleFrom = vbPoints) And (ScaleTo = vbPixels)
        ConvertScale = sngValue * 20 / sngTPP_X
    Case (ScaleFrom = vbPixels) And (ScaleTo = vbHimetric)
        ConvertScale = sngValue * HimetricPerPixel
    Case (ScaleFrom = vbHimetric) And (ScaleTo = vbPixels)
        ConvertScale = sngValue / HimetricPerPixel
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ConvertScale")
Resume out:
End Function

'Public Function ScreenWidth(Optional ByVal Actual As Boolean) As Single
'    ' returned as Twips
'    If Actual Then
'        Dim hDC As Long: hDC = GetDC(0)
'        ScreenWidth = GetDeviceCaps(hDC, 118) * Screen.TwipsPerPixelX ' 118=DESKTOPHORZRES
'        ReleaseDC 0, hDC
'    Else
'        ScreenWidth = Screen.Width
'    End If
'
'End Function
'
'Public Function ScreenHeight(Optional ByVal Actual As Boolean) As Single
'    ' returned as Twips
'    If Actual Then
'        Dim hDC As Long: hDC = GetDC(0)
'        ScreenHeight = GetDeviceCaps(hDC, 117) * Screen.TwipsPerPixelY ' 117=DESKTOPVERTRES
'        ReleaseDC 0, hDC
'    Else
'        ScreenHeight = Screen.Height
'    End If
'End Function
'
'Public Function ScreenDPI(Optional ByVal Actual As Boolean) As Single
'    If Actual Then
'        Dim hDC As Long: hDC = GetDC(0)
'        ScreenDPI = GetDeviceCaps(hDC, 118) / (Screen.Width / Screen.TwipsPerPixelX)
'        ReleaseDC 0, hDC
'        If ScreenDPI = 1 Then
'            ScreenDPI = 1440! / Screen.TwipsPerPixelX
'        Else
'            ScreenDPI = ScreenDPI * 96!
'        End If
'    Else
'        ScreenDPI = 1440! / Screen.TwipsPerPixelX
'    End If
'End Function

